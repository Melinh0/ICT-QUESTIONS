#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
Bot para obter gabaritos e explicações do Copilot para questões de concurso.

Uso:
    python3 obter_gabaritos_copilot.py --input questoes_organizadas.json --output questoes_com_gabaritos.json

Recomendações:
- Execute com o ambiente virtual ativado (.venv).
- Posicione o mouse sobre o campo do Copilot antes do envio.
- Ajuste coordenadas se necessário.
"""
import os
import sys
import json
import time
import re
import argparse
import pyautogui
import pyperclip

# Coordenadas padrão (ajuste conforme necessário)
COORD = {
    "campo_prompt": (3218, 559),
    "enviar_mensagem": (3790, 717),
    "copiar_resposta": (3318, 897),
    "ver_mais_resposta": (3500, 895),
    "icone_chat": (3144, 535),
    "excluir_chat": (3203, 690),
    "confirmar_exclusao": (3368, 622)
}

WAIT_AFTER_SEND = 45  # Tempo de espera para resposta
MAX_RETRIES = 3  # Máximo de tentativas por lote
QUESTIONS_PER_BATCH = 5  # Questões por lote

def verificar_dependencias():
    try:
        import pyautogui, pyperclip  # noqa: F401
        return True
    except ImportError as e:
        print("Dependências ausentes:", e)
        print("Instale: pip install pyautogui pyperclip")
        return False

def carregar_questoes(path):
    """Carrega as questões do arquivo JSON"""
    with open(path, "r", encoding="utf-8") as f:
        data = json.load(f)
    return data

def preparar_questoes(questoes_originais):
    """Prepara as questões para processamento"""
    questoes_preparadas = []
    
    for idx, questao in enumerate(questoes_originais):
        questao_preparada = {
            "numero": idx + 1,
            "codigo": questao.get("codigo", ""),
            "banca": questao.get("banca", ""),
            "tema": questao.get("tema", ""),
            "assunto": questao.get("assunto", ""),
            "concurso": questao.get("concurso", ""),
            "cargo": questao.get("cargo", ""),
            "enunciado": questao.get("enunciado", ""),
            "alternativas": questao.get("alternativas", {}),
            "gabarito_copilot": "",  # Será preenchido pelo Copilot
            "explicacao_copilot": ""  # Será preenchido pelo Copilot
        }
        questoes_preparadas.append(questao_preparada)
    
    return questoes_preparadas

def clicar(coordenada, dur=0.5, duplo=False):
    x, y = coordenada
    pyautogui.moveTo(x, y, duration=dur)
    time.sleep(0.3)
    if duplo:
        pyautogui.doubleClick()
    else:
        pyautogui.click()
    time.sleep(0.3)

def limpar_chat():
    """Limpa o chat do Copilot para nova conversa"""
    try:
        clicar(COORD["icone_chat"])
        time.sleep(0.8)
        clicar(COORD["excluir_chat"])
        time.sleep(0.8)
        clicar(COORD["confirmar_exclusao"])
        time.sleep(1.5)
    except Exception:
        pass

def enviar_para_copilot(texto):
    """Cola o texto no campo do Copilot, envia e retorna resposta"""
    try:
        clicar(COORD["campo_prompt"])
        pyperclip.copy(texto)
        time.sleep(0.4)
        pyautogui.hotkey("ctrl", "v")
        time.sleep(0.4)
        clicar(COORD["enviar_mensagem"])
        print("  ⏳ Aguardando resposta do Copilot...")
        time.sleep(WAIT_AFTER_SEND)
        
        # Tentar ver mais se disponível
        try:
            clicar(COORD["ver_mais_resposta"])
            time.sleep(2)
        except Exception:
            pass
            
        # Copiar resposta
        clicar(COORD["copiar_resposta"])
        time.sleep(1.2)
        return pyperclip.paste()
    except Exception as e:
        print("Erro ao enviar para Copilot:", e)
        return ""

def montar_prompt_lote(questoes_lote):
    """Monta o prompt para um lote de questões"""
    num_questoes = len(questoes_lote)
    
    prompt = (
        f"Para cada uma das {num_questoes} questões de concurso abaixo, forneça:\n"
        "1. O GABARITO CORRETO (apenas a letra da alternativa)\n"
        "2. Uma EXPLICAÇÃO detalhada e bem fundamentada sobre por que essa alternativa está correta\n\n"
        "REQUISITOS:\n"
        "- Baseie-se em fontes confiáveis e na legislação/doutrina pertinente\n"
        "- A explicação deve ser técnica e embasada\n"
        "- Considere a banca examinadora (FGV) e o contexto do concurso\n\n"
    )
    
    for i, questao in enumerate(questoes_lote, 1):
        prompt += f"QUESTÃO {i}:\n"
        prompt += f"Código: {questao['codigo']}\n"
        prompt += f"Banca: {questao['banca']}\n"
        prompt += f"Concurso: {questao['concurso']}\n"
        prompt += f"Cargo: {questao.get('cargo', 'Não informado')}\n"
        prompt += f"Tema: {questao['tema']}\n"
        prompt += f"Assunto: {questao['assunto']}\n"
        prompt += f"ENUNCIADO:\n{questao['enunciado']}\n\n"
        prompt += "ALTERNATIVAS:\n"
        
        for letra, texto in sorted(questao['alternativas'].items()):
            prompt += f"{letra}) {texto}\n"
        
        prompt += "\n" + "-" * 50 + "\n\n"
    
    prompt += (
        "FORMATO DE RESPOSTA EXATO (obrigatório):\n"
        "Para cada questão, use exatamente:\n"
        f"##### QUESTÃO X #####\n"
        "GABARITO: [A, B, C, D ou E]\n"
        "EXPLICAÇÃO: [sua explicação completa e fundamentada aqui]\n"
        f"##### FIM QUESTÃO X #####\n\n"
        f"Substitua X pelo número da questão (1 a {num_questoes}).\n"
        "Não inclua qualquer outro texto fora deste formato."
    )
    
    return prompt

def extrair_gabarito_explicacao(texto_resposta, num_questoes):
    """
    Extrai gabarito e explicação para todas as questões do lote.
    Retorna lista de tuplas (gabarito, explicacao) na mesma ordem das questões.
    """
    resultados = [("", "")] * num_questoes
    
    # Padrão para extração
    for i in range(1, num_questoes + 1):
        padrao = f'##### QUESTÃO {i} #####(.*?)##### FIM QUESTÃO {i} #####'
        match = re.search(padrao, texto_resposta, re.DOTALL | re.IGNORECASE)
        
        if match:
            conteudo = match.group(1).strip()
            
            # Extrair gabarito
            gabarito_match = re.search(r'GABARITO:\s*([A-E])', conteudo, re.IGNORECASE)
            gabarito = gabarito_match.group(1).upper() if gabarito_match else ""
            
            # Extrair explicação
            explicacao_match = re.search(r'EXPLICAÇÃO:\s*(.*?)(?=$|GABARITO:)', conteudo, re.DOTALL | re.IGNORECASE)
            explicacao = explicacao_match.group(1).strip() if explicacao_match else ""
            
            resultados[i-1] = (gabarito, explicacao)
    
    # Fallback: tentar padrões mais flexíveis se não encontrou no formato exato
    if any(not gab or not exp for gab, exp in resultados):
        for i in range(1, num_questoes + 1):
            if not resultados[i-1][0] or not resultados[i-1][1]:
                # Procurar por padrões alternativos
                padroes_flexiveis = [
                    f'QUESTÃO {i}.*?GABARITO[\\s:]*([A-E]).*?EXPLICAÇÃO[\\s:]*(.*?)(?=QUESTÃO|#####|$)',
                    f'Questão {i}.*?Gabarito[\\s:]*([A-E]).*?Explicação[\\s:]*(.*?)(?=Questão|#####|$)',
                ]
                
                for padrao in padroes_flexiveis:
                    match = re.search(padrao, texto_resposta, re.DOTALL | re.IGNORECASE)
                    if match:
                        gabarito = match.group(1).upper() if match.group(1) else ""
                        explicacao = match.group(2).strip() if match.group(2) else ""
                        if gabarito and len(explicacao) > 50:
                            resultados[i-1] = (gabarito, explicacao)
                            break
    
    return resultados

def processar_lote_questoes(questoes_lote, tentativa=1):
    """
    Processa um lote de questões com retry logic.
    Retorna lista de (gabarito, explicacao) para o lote.
    """
    num_questoes = len(questoes_lote)
    print(f"  📝 Processando lote de {num_questoes} questões (tentativa {tentativa})...")
    print(f"  Questões no lote: {[q['codigo'] for q in questoes_lote]}")
    
    prompt = montar_prompt_lote(questoes_lote)
    resposta = enviar_para_copilot(prompt)
    
    if not resposta:
        print(f"  ⚠️ Resposta vazia para o lote")
        return None
    
    resultados = extrair_gabarito_explicacao(resposta, num_questoes)
    
    # Verificar quantos resultados válidos foram obtidos
    resultados_validos = sum(1 for gab, exp in resultados if gab and exp)
    print(f"  📊 Resultados obtidos: {resultados_validos}/{num_questoes}")
    
    if resultados_validos > 0:
        return resultados
    else:
        print(f"  ❌ Não foi possível extrair resultados da resposta")
        # Salvar resposta bruta para debug
        with open(f"debug_resposta_tentativa_{tentativa}.txt", "w", encoding="utf-8") as f:
            f.write(f"PROMPT:\n{prompt}\n\nRESPOSTA:\n{resposta}")
        return None

def salvar_progresso(questoes, output_path):
    """Salva as questões com gabaritos e explicações"""
    try:
        with open(output_path, "w", encoding="utf-8") as f:
            json.dump(questoes, f, ensure_ascii=False, indent=2)
        return True
    except Exception as e:
        print(f"  ❌ Erro ao salvar arquivo: {e}")
        return False

def processar_todas_questoes(input_path, output_path):
    """Processa todas as questões do arquivo JSON"""
    print(f"📖 Carregando questões de: {input_path}")
    questoes_originais = carregar_questoes(input_path)
    questoes = preparar_questoes(questoes_originais)
    total = len(questoes)
    
    print(f"📊 Total de questões detectadas: {total}")
    
    if total == 0:
        print("❌ Nenhuma questão detectada!")
        return

    # Verificar progresso existente
    completed_codes = set()
    if os.path.exists(output_path):
        try:
            with open(output_path, "r", encoding="utf-8") as f:
                existente = json.load(f)
            
            for questao in existente:
                codigo = questao.get("codigo")
                gabarito = questao.get("gabarito_copilot", "")
                explicacao = questao.get("explicacao_copilot", "")
                if codigo and gabarito and explicacao:
                    completed_codes.add(codigo)
            
            print(f"🔄 {len(completed_codes)} questões já processadas. Continuando...")
        except Exception as e:
            print(f"⚠️ Erro ao ler arquivo existente: {e}")

    # Filtrar questões pendentes
    pendentes = [q for q in questoes if q["codigo"] not in completed_codes]
    
    if not pendentes:
        print("✅ Todas as questões já foram processadas!")
        return

    print(f"📝 Questões pendentes: {len(pendentes)}")
    
    # Dividir em lotes
    lotes = []
    for i in range(0, len(pendentes), QUESTIONS_PER_BATCH):
        lote = pendentes[i:i + QUESTIONS_PER_BATCH]
        lotes.append(lote)
    
    print(f"📦 Total de lotes a processar: {len(lotes)}")
    print(f"🔢 Tamanho dos lotes: {QUESTIONS_PER_BATCH} questões")
    
    # Contador regressivo para posicionamento
    print("🎯 Posicione o mouse no campo do Copilot. Iniciando em 10 segundos...")
    for i in range(10, 0, -1):
        print(f"  {i}...")
        time.sleep(1)

    # Processar cada lote
    for lote_num, lote_questoes in enumerate(lotes, 1):
        num_questoes_lote = len(lote_questoes)
        print(f"\n{'='*60}")
        print(f"--- Lote {lote_num}/{len(lotes)} ({num_questoes_lote} questões) ---")
        print(f"{'='*60}")
        
        resultados = None
        for tentativa in range(1, MAX_RETRIES + 1):
            limpar_chat()
            time.sleep(2)
            
            resultados = processar_lote_questoes(lote_questoes, tentativa)
            
            if resultados:
                break
            elif tentativa < MAX_RETRIES:
                print(f"  🔄 Tentando novamente... ({tentativa + 1}/{MAX_RETRIES})")
                time.sleep(5)
        
        # Atribuir resultados às questões
        if resultados:
            for i, questao in enumerate(lote_questoes):
                if i < len(resultados):
                    gabarito, explicacao = resultados[i]
                    if gabarito and explicacao:
                        questao["gabarito_copilot"] = gabarito
                        questao["explicacao_copilot"] = explicacao
                        print(f"  ✅ {questao['codigo']}: Gabarito '{gabarito}' e explicação obtida")
                    else:
                        questao["gabarito_copilot"] = "[ERRO: Gabarito não identificado]"
                        questao["explicacao_copilot"] = "[ERRO: Explicação não obtida]"
                        print(f"  ⚠️ {questao['codigo']}: Resultados incompletos")
                else:
                    questao["gabarito_copilot"] = "[ERRO: Fora do range]"
                    questao["explicacao_copilot"] = "[ERRO: Fora do range]"
                    print(f"  ❌ {questao['codigo']}: Índice fora do range")
        else:
            for questao in lote_questoes:
                questao["gabarito_copilot"] = "[ERRO: Falha após múltiplas tentativas]"
                questao["explicacao_copilot"] = "[ERRO: Falha após múltiplas tentativas]"
            print(f"  💀 Falha crítica no lote {lote_num}")
        
        # Salvar progresso após cada lote
        if salvar_progresso(questoes, output_path):
            print(f"  💾 Progresso salvo: Lote {lote_num}/{len(lotes)} completo")
        else:
            print(f"  ❌ Falha ao salvar progresso do lote {lote_num}")
        
        # Pausa entre lotes
        if lote_num < len(lotes):
            print("  ⏳ Aguardando 10 segundos para próximo lote...")
            time.sleep(10)

    print(f"\n{'='*60}")
    print(f"✅ Processo finalizado! Arquivo salvo em: {output_path}")
    print(f"📊 Total de questões processadas: {total}")
    print(f"{'='*60}")

def main():
    parser = argparse.ArgumentParser(description="Obter gabaritos e explicações do Copilot para questões de concurso")
    parser.add_argument("--input", "-i", required=True, help="Arquivo JSON de entrada com as questões")
    parser.add_argument("--output", "-o", required=True, help="Arquivo JSON de saída com gabaritos e explicações")
    
    args = parser.parse_args()

    if not verificar_dependencias():
        sys.exit(1)

    if not os.path.exists(args.input):
        print("❌ Arquivo de entrada não encontrado:", args.input)
        sys.exit(1)

    # Definir pasta de saída
    output_dir = "questões_CPRM"
    output_path = os.path.join(output_dir, args.output) if not os.path.isabs(args.output) else args.output
    
    # Criar diretório se não existir
    os.makedirs(os.path.dirname(output_path), exist_ok=True)
    
    processar_todas_questoes(args.input, output_path)

if __name__ == "__main__":
    main()
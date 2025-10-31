#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
Gera explicações para 5 questões por vez usando automação do Copilot.

Uso:
    python3 gerar_explicacoes_copilot.py --input path/to/questions_fixture_unificado.json \
                                        --output explicacoes_completas.json

Recomendações:
- Execute com o ambiente virtual ativado (.venv).
- Posicione o mouse sobre o campo do Copilot antes do envio (há contadores).
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

# Coordenadas padrão (copie dos seus scripts se precisar ajustar)
COORD = {
    "campo_prompt": (3218, 559),
    "enviar_mensagem": (3790, 717),
    "copiar_resposta": (3318, 897),
    "ver_mais_resposta": (3500, 895),
    "icone_chat": (3144, 535),
    "excluir_chat": (3203, 690),
    "confirmar_exclusao": (3368, 622)
}

WAIT_AFTER_SEND = 45  # Reduzido para 5 questões
MAX_RETRIES = 3  # Máximo de tentativas por lote
QUESTIONS_PER_BATCH = 5  # 695 ÷ 5 = 139 lotes exatos


def verificar_dependencias():
    try:
        import pyautogui, pyperclip  # noqa: F401
        return True
    except ImportError as e:
        print("Dependências ausentes:", e)
        print("Instale: pip install pyautogui pyperclip")
        return False


def carregar_fixture(path):
    with open(path, "r", encoding="utf-8") as f:
        data = json.load(f)
    return data


def reconstruir_questoes(fixt):
    """
    Reconstrói estrutura de questões a partir do JSON de fixtures.
    Retorna lista de dicts: {numero: pk_question, enunciado, itens: {A: text...}, item_correto, explicacao: ""}
    """
    questions = {}
    alternatives_by_question = {}
    
    print(f"Debug: Processando {len(fixt)} objetos no total")
    
    # Primeiro passe: coletar todas as questões
    for obj in fixt:
        model = obj.get("model", "")
        pk = obj.get("pk")
        fields = obj.get("fields", {}) or {}

        if model.endswith(".Question"):
            questions[pk] = {
                "pk": pk,
                "enunciado": fields.get("text", "").strip(),
                "alternatives": [],
                "item_correto": None,
                "explicacao": "",
            }
            alternatives_by_question[pk] = []

    print(f"Debug: Encontradas {len(questions)} questões")

    # Segundo passe: coletar alternativas
    for obj in fixt:
        model = obj.get("model", "")
        pk = obj.get("pk")
        fields = obj.get("fields", {}) or {}

        if model.endswith(".Alternative"):
            question_id = fields.get("question")
            if question_id in questions:
                alternative_data = {
                    "pk": pk,
                    "text": fields.get("text", "").strip(),
                    "is_correct": fields.get("is_correct", False)
                }
                alternatives_by_question[question_id].append(alternative_data)
                
                if fields.get("is_correct") and questions[question_id]["item_correto"] is None:
                    questions[question_id]["item_correto"] = pk

    # Organizar alternativas por PK e determinar letras
    lista = []
    for question_pk in sorted(questions.keys()):
        q = questions[question_pk]
        
        alternatives = sorted(alternatives_by_question.get(question_pk, []), key=lambda x: x["pk"])
        
        itens = {}
        correct_letter = None
        
        for idx, alt in enumerate(alternatives):
            letra = chr(ord("A") + idx)
            itens[letra] = alt["text"]
            
            if alt["pk"] == q["item_correto"]:
                correct_letter = letra
        
        lista.append({
            "numero": len(lista) + 1,
            "pk": question_pk,
            "enunciado": q["enunciado"],
            "itens": itens,
            "item_correto": correct_letter,
            "explicacao": ""
        })
    
    print(f"Debug: Reconstruídas {len(lista)} questões completas")
    return lista


def converter_para_fixture_final(questoes_com_explicacoes):
    """
    Converte a estrutura interna de volta para o formato de fixture Django
    com os models corretos: "questions.Question" e "alternatives.Alternative"
    """
    fixture_final = []
    
    for questao in questoes_com_explicacoes:
        fixture_final.append({
            "model": "questions.Question",
            "pk": questao["pk"],
            "fields": {
                "text": questao["enunciado"],
                "explicacao": questao["explicacao"],
                "level": "HCIA",
                "track": "Computing",
                "has_answer": True,
                "has_multiple_answers": False,
                "weight": "1.00"
            }
        })
        
        for idx, (letra, texto) in enumerate(sorted(questao["itens"].items())):
            is_correct = (letra == questao["item_correto"])
            alternative_pk = questao["pk"] * 10 + idx
            
            fixture_final.append({
                "model": "alternatives.Alternative",
                "pk": alternative_pk,
                "fields": {
                    "question": questao["pk"],
                    "text": texto,
                    "is_correct": is_correct
                }
            })
    
    return fixture_final


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
    """Cola o texto no campo do Copilot, envia e retorna texto da área de transferência"""
    try:
        clicar(COORD["campo_prompt"])
        pyperclip.copy(texto)
        time.sleep(0.4)
        pyautogui.hotkey("ctrl", "v")
        time.sleep(0.4)
        clicar(COORD["enviar_mensagem"])
        print("  ⏳ aguardando resposta do Copilot...")
        time.sleep(WAIT_AFTER_SEND)
        try:
            clicar(COORD["ver_mais_resposta"])
            time.sleep(2)
        except Exception:
            pass
        clicar(COORD["copiar_resposta"])
        time.sleep(1.2)
        return pyperclip.paste()
    except Exception as e:
        print("Erro ao enviar para Copilot:", e)
        return ""


def montar_prompt_lote(questoes_lote):
    """
    Monta o prompt para um lote de questões.
    """
    num_questoes = len(questoes_lote)
    prompt = (
        f"Para cada uma das {num_questoes} questões abaixo, forneça APENAS a EXPLICAÇÃO do porquê o ITEM CORRETO está correto.\n\n"
    )
    
    for i, questao in enumerate(questoes_lote, 1):
        prompt += f"QUESTÃO {i}:\n"
        prompt += f"ENUNCIADO:\n{questao['enunciado']}\n\n"
        prompt += f"ITENS:\n"
        
        for letra, texto in sorted(questao['itens'].items()):
            prompt += f"{letra}) {texto}\n"
        
        prompt += f"ITEM CORRETO: {questao['item_correto'] or 'N/A'}\n\n"
        prompt += "---\n\n"
    
    prompt += (
        "FORMATO DE RESPOSTA EXATO (obrigatório):\n"
        "Para cada questão, use exatamente:\n"
        f"##### EXPLICAÇÃO QUESTÃO X #####\n"
        "[sua explicação completa aqui]\n"
        f"##### FIM EXPLICAÇÃO QUESTÃO X #####\n\n"
        f"Substitua X pelo número da questão (1 a {num_questoes})."
    )
    
    return prompt


def extrair_explicacoes_lote(texto_resposta, num_questoes):
    """
    Extrai explicações para todas as questões do lote.
    Retorna lista de explicações na mesma ordem das questões.
    """
    explicacoes = [""] * num_questoes
    
    # Tentar padrão exato primeiro
    for i in range(1, num_questoes + 1):
        padrao = f'##### EXPLICAÇÃO QUESTÃO {i} #####(.*?)##### FIM EXPLICAÇÃO QUESTÃO {i} #####'
        match = re.search(padrao, texto_resposta, re.DOTALL | re.IGNORECASE)
        if match:
            explicacoes[i-1] = match.group(1).strip()
    
    # Se não encontrou todas, tentar padrões mais flexíveis
    if any(not exp for exp in explicacoes):
        for i in range(1, num_questoes + 1):
            if not explicacoes[i-1]:  # Ainda não tem explicação para esta
                padroes_flexiveis = [
                    f'EXPLICAÇÃO QUESTÃO {i}[:\\s]*(.*?)(?=#####|QUESTÃO|EXPLICAÇÃO|$)',
                    f'QUESTÃO {i}.*?EXPLICAÇÃO[:\\s]*(.*?)(?=QUESTÃO|#####|$)',
                    f'Questão {i}.*?Explicação[:\\s]*(.*?)(?=Questão|#####|$)',
                ]
                
                for padrao in padroes_flexiveis:
                    match = re.search(padrao, texto_resposta, re.DOTALL | re.IGNORECASE)
                    if match:
                        explicacao = match.group(1).strip()
                        explicacao = re.sub(r'^[:-\s]*', '', explicacao)
                        if len(explicacao) > 50:
                            explicacoes[i-1] = explicacao
                            break
    
    return explicacoes


def processar_lote_questoes(questoes_lote, tentativa=1):
    """
    Processa um lote de questões com retry logic.
    """
    num_questoes = len(questoes_lote)
    print(f"  📝 Processando lote de {num_questoes} questões (tentativa {tentativa})...")
    print(f"  Questões no lote: {[q['numero'] for q in questoes_lote]}")
    
    prompt = montar_prompt_lote(questoes_lote)
    resposta = enviar_para_copilot(prompt)
    
    if not resposta:
        print(f"  ⚠️ Resposta vazia para o lote")
        return None
    
    explicacoes = extrair_explicacoes_lote(resposta, num_questoes)
    
    # Verificar quantas explicações foram obtidas
    explicacoes_obtidas = sum(1 for exp in explicacoes if exp)
    print(f"  📊 Explicações obtidas: {explicacoes_obtidas}/{num_questoes}")
    
    if explicacoes_obtidas > 0:
        return explicacoes
    else:
        print(f"  ❌ Não foi possível extrair explicações da resposta")
        # Salvar resposta bruta para debug
        with open(f"debug_lote_tentativa_{tentativa}.txt", "w", encoding="utf-8") as f:
            f.write(f"PROMPT:\n{prompt}\n\nRESPOSTA:\n{resposta}")
        return None


def processar_todos(fixtures_path, out_path):
    fixt = carregar_fixture(fixtures_path)
    questoes = reconstruir_questoes(fixt)
    total = len(questoes)
    print(f"Questões detectadas: {total}")

    if total == 0:
        print("❌ Nenhuma questão detectada!")
        return

    # Verificar progresso existente
    completed_nums = set()
    if os.path.exists(out_path):
        try:
            with open(out_path, "r", encoding="utf-8") as f:
                existente = json.load(f)
            
            if existente and isinstance(existente[0], dict) and "model" in existente[0]:
                temp_questoes = reconstruir_questoes(existente)
                for q in temp_questoes:
                    if q.get("explicacao", "").strip():
                        completed_nums.add(q["numero"])
            else:
                for q in existente:
                    if isinstance(q, dict):
                        num = q.get("numero")
                        explic = q.get("explicacao", "")
                        if num and explic.strip():
                            completed_nums.add(num)
            
            print(f"[RESUME] {len(completed_nums)} questões já processadas. Continuando...")
        except Exception as e:
            print(f"[WARN] Erro ao ler arquivo existente: {e}")

    # Filtrar questões pendentes
    pendentes = [q for q in questoes if q["numero"] not in completed_nums]
    if not pendentes:
        print("Todas as questões já foram processadas!")
        return

    print(f"Questões pendentes: {len(pendentes)}")
    
    # Dividir em lotes de 5 questões (divisor de 695)
    lotes = []
    for i in range(0, len(pendentes), QUESTIONS_PER_BATCH):
        lote = pendentes[i:i + QUESTIONS_PER_BATCH]
        lotes.append(lote)
    
    print(f"Total de lotes a processar: {len(lotes)}")
    print(f"Tamanho dos lotes: {QUESTIONS_PER_BATCH} questões")
    
    # Contador regressivo para posicionamento
    print("Posicione o mouse no campo do Copilot. Iniciando em 10 segundos...")
    for i in range(10, 0, -1):
        print(f"  {i}...")
        time.sleep(1)

    # Processar cada lote
    for lote_num, lote_questoes in enumerate(lotes, 1):
        num_questoes_lote = len(lote_questoes)
        print(f"\n{'='*50}")
        print(f"--- Lote {lote_num}/{len(lotes)} ({num_questoes_lote} questões) ---")
        print(f"{'='*50}")
        
        explicacoes = None
        for tentativa in range(1, MAX_RETRIES + 1):
            limpar_chat()
            time.sleep(2)
            
            explicacoes = processar_lote_questoes(lote_questoes, tentativa)
            
            if explicacoes:
                break
            elif tentativa < MAX_RETRIES:
                print(f"  🔄 Tentando novamente... ({tentativa + 1}/{MAX_RETRIES})")
                time.sleep(5)
        
        # Atribuir explicações às questões
        if explicacoes:
            for i, questao in enumerate(lote_questoes):
                if i < len(explicacoes) and explicacoes[i]:
                    questao["explicacao"] = explicacoes[i]
                    print(f"  ✅ Q{questao['numero']}: Explicação obtida ({len(explicacoes[i])} chars)")
                else:
                    questao["explicacao"] = "[ERRO: Não foi possível obter explicação para esta questão]"
                    print(f"  ❌ Q{questao['numero']}: Falha na explicação")
        else:
            for questao in lote_questoes:
                questao["explicacao"] = "[ERRO: Não foi possível obter explicação após múltiplas tentativas]"
            print(f"  💀 Falha crítica no lote {lote_num}")
        
        # Salvar progresso após cada lote
        try:
            fixture_final = converter_para_fixture_final(questoes)
            with open(out_path, "w", encoding="utf-8") as f:
                json.dump(fixture_final, f, ensure_ascii=False, indent=2)
            print(f"  💾 Progresso salvo: Lote {lote_num}/{len(lotes)} completo")
        except Exception as e:
            print(f"  [ERRO] Ao salvar: {e}")
        
        # Pausa entre lotes
        if lote_num < len(lotes):
            print("  ⏳ Aguardando 8 segundos para próximo lote...")
            time.sleep(8)

    print(f"\n{'='*50}")
    print(f"✅ Processo finalizado! Arquivo salvo em: {out_path}")
    print(f"📊 Total de questões processadas: {total}")
    print(f"{'='*50}")


def main():
    parser = argparse.ArgumentParser(description="Gerar explicações com Copilot a partir de fixture JSON")
    parser.add_argument("--input", "-i", required=True, help="arquivo JSON de entrada (fixture unificado)")
    parser.add_argument("--output", "-o", required=True, help="arquivo JSON de saída com explicações")
    args = parser.parse_args()

    if not verificar_dependencias():
        sys.exit(1)

    if not os.path.exists(args.input):
        print("Arquivo de entrada não encontrado:", args.input)
        sys.exit(1)

    processar_todos(args.input, args.output)


if __name__ == "__main__":
    main()
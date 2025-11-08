#!/usr/bin/env python3
# -*- coding: utf-8 -*-

import os
import re
import json
import time
import pyautogui
import pyperclip
from datetime import datetime
from docx import Document
import xml.etree.ElementTree as ET

class ExtratorQuestoesDOCX:
    def __init__(self):
        self.coordenadas = {
            "campo_prompt": (3218, 559),
            "enviar_mensagem": (3790, 717),
            "copiar_resposta": (3318, 897),
            "ver_mais_resposta": (3500, 895),
            "icone_chat": (3144, 535),
            "excluir_chat": (3203, 690),
            "confirmar_exclusao": (3368, 622)
        }
        
        # Prompt para organizar as questões - MAIS ESPECÍFICO
        self.prompt_organizacao = """ORGANIZE E SEPARE as seguintes questões de múltipla escolha. 
Para CADA questão identificada, formate EXATAMENTE assim:

##### QUESTÃO [NÚMERO] #####
ENUNCIADO: [texto completo do enunciado APENAS - SEM RESPOSTAS]
ITENS:
A) [texto completo da alternativa A]
B) [texto completo da alternativa B] 
C) [texto completo da alternativa C]
D) [texto completo da alternativa D]
E) [texto completo da alternativa E] (se existir)
ITEM CORRETO: [letra da alternativa correta se disponível no texto]
EXPLICAÇÃO: [explicação oficial se disponível]
##### FIM QUESTÃO [NÚMERO] #####

CRÍTICO: 
- O ENUNCIADO deve conter APENAS a pergunta original, SEM respostas ou explicações
- NÃO inclua respostas do Copilot no enunciado
- Se encontrar explicações ou respostas no texto do enunciado, REMOVA-AS
- Mantenha o formato exato acima.

Aqui estão as questões para organizar:"""

        # Prompt para obter gabarito - MAIS DIRETO
        self.prompt_gabarito = """Analise esta questão de múltipla escolha e forneça o gabarito mais provável baseado no conteúdo técnico. 

ENUNCIADO: [texto completo do enunciado]

ITENS:
A) [texto completo da alternativa A]
B) [texto completo da alternativa B]
C) [texto completo da alternativa C] 
D) [texto completo da alternativa D]
E) [texto completo da alternativa E] (se existir)

Formate a resposta EXATAMENTE assim:

ITEM CORRETO: [letra da alternativa que você considera correta]

EXPLICAÇÃO: [sua análise técnica detalhada explicando por que esta é a resposta correta]

NÃO inclua o enunciado ou alternativas na resposta."""
    
    def verificar_dependencias(self):
        """Verifica se todas as dependências estão disponíveis"""
        try:
            import pyautogui
            import pyperclip
            from docx import Document
            print("✅ Todas as dependências encontradas")
            return True
        except ImportError as e:
            print(f"❌ Dependências não encontradas: {e}")
            print("Instale com: pip install pyautogui pyperclip python-docx")
            return False
    
    def extrair_texto_docx(self, arquivo_docx):
        """Extrai todo o texto do documento DOCX"""
        print(f"📖 Extraindo texto de: {arquivo_docx}")
        
        try:
            doc = Document(arquivo_docx)
            texto_completo = ""
            
            for paragraph in doc.paragraphs:
                texto_completo += paragraph.text + "\n"
            
            print(f"✅ Texto extraído: {len(texto_completo)} caracteres")
            return texto_completo
            
        except Exception as e:
            print(f"❌ Erro ao extrair texto do DOCX: {e}")
            return ""

    def dividir_texto_em_partes(self, texto, max_caracteres=2500):
        """Divide o texto em partes menores para envio ao Copilot"""
        print(f"📝 Dividindo texto em partes de até {max_caracteres} caracteres...")
        
        partes = []
        linhas = texto.split('\n')
        parte_atual = ""
        
        for linha in linhas:
            # Se adicionar esta linha não ultrapassar o limite, adiciona à parte atual
            if len(parte_atual) + len(linha) + 1 <= max_caracteres:
                parte_atual += linha + "\n"
            else:
                # Se a parte atual não está vazia, adiciona à lista
                if parte_atual.strip():
                    partes.append(parte_atual.strip())
                # Começa nova parte com a linha atual
                parte_atual = linha + "\n"
        
        # Adiciona a última parte se não estiver vazia
        if parte_atual.strip():
            partes.append(parte_atual.strip())
        
        print(f"✅ Texto dividido em {len(partes)} partes")
        for i, parte in enumerate(partes):
            print(f"   Parte {i+1}: {len(parte)} caracteres")
        
        return partes

    def clicar_coordenada(self, coordenada, descricao=""):
        """Clica em uma coordenada específica"""
        print(f"  🖱️ Clicando em {descricao}: {coordenada}")
        x, y = coordenada
        pyautogui.moveTo(x, y, duration=0.5)
        time.sleep(0.5)
        pyautogui.click()
        time.sleep(1)
    
    def colar_texto(self, texto, descricao=""):
        """Copia e cola texto"""
        print(f"  📋 Colando {descricao} ({len(texto)} caracteres)")
        pyperclip.copy(texto)
        time.sleep(0.5)
        pyautogui.hotkey('ctrl', 'v')
        time.sleep(1)
    
    def limpar_chat(self):
        """Limpa o chat do Copilot"""
        try:
            self.clicar_coordenada(
                self.coordenadas['icone_chat'], 
                "Ícone do chat"
            )
            time.sleep(1)
            
            self.clicar_coordenada(
                self.coordenadas['excluir_chat'], 
                "Excluir chat"
            )
            time.sleep(1)
            
            self.clicar_coordenada(
                self.coordenadas['confirmar_exclusao'], 
                "Confirmar exclusão"
            )
            time.sleep(2)
        except Exception as e:
            print(f"  ⚠️ Erro ao limpar chat: {e}")
    
    def ver_mais_resposta(self):
        """Clica no botão 'Ver mais' para carregar a resposta completa"""
        print("  🔍 Clicando em 'Ver mais' para carregar resposta completa...")
        try:
            self.clicar_coordenada(
                self.coordenadas['ver_mais_resposta'], 
                "Ver mais resposta"
            )
            time.sleep(3)
            return True
        except Exception as e:
            print(f"  ⚠️ Erro ao clicar em 'Ver mais': {e}")
            return False
    
    def enviar_para_copilot(self, texto, descricao=""):
        """Envia texto para o Copilot e retorna a resposta"""
        print(f"\n🔄 ENVIANDO {descricao} PARA COPILOT")
        
        try:
            # 1. Clicar no campo de prompt
            self.clicar_coordenada(
                self.coordenadas['campo_prompt'], 
                "Campo de prompt"
            )
            
            # 2. Colar o texto
            self.colar_texto(texto, descricao)
            
            # 3. Enviar mensagem
            self.clicar_coordenada(
                self.coordenadas['enviar_mensagem'], 
                "Enviar mensagem"
            )
            
            print("  ⏳ Aguardando resposta do Copilot (35 segundos)...")
            time.sleep(35)
            
            # 4. Clicar em "Ver mais" para carregar resposta completa
            self.ver_mais_resposta()
            
            # 5. Copiar resposta
            self.clicar_coordenada(
                self.coordenadas['copiar_resposta'], 
                "Copiar resposta"
            )
            time.sleep(2)
            
            # 6. Obter texto da área de transferência
            try:
                texto_resposta = pyperclip.paste()
                print(f"  ✅ Resposta copiada ({len(texto_resposta)} caracteres)")
                
                return texto_resposta
                
            except Exception as e:
                print(f"  ❌ Erro ao acessar área de transferência: {e}")
                return ""
            
        except Exception as e:
            print(f"❌ Erro ao enviar para Copilot: {e}")
            return ""

    def organizar_questoes_com_copilot(self, texto_completo):
        """Usa o Copilot para organizar e separar todas as questões"""
        print("🎯 ORGANIZANDO QUESTÕES COM COPILOT")
        
        # Dividir o texto em partes menores
        partes_texto = self.dividir_texto_em_partes(texto_completo, max_caracteres=2000)
        
        todas_questoes_organizadas = []
        
        # Contador regressivo
        print("⏰ Iniciando em 5 segundos... Posicione o mouse no campo do Copilot!")
        for i in range(5, 0, -1):
            print(f"  {i}...")
            time.sleep(1)
        
        for i, parte in enumerate(partes_texto, 1):
            print(f"\n{'='*50}")
            print(f"📦 PROCESSANDO PARTE {i}/{len(partes_texto)}")
            print(f"{'='*50}")
            
            # Preparar texto para envio
            texto_para_envio = f"{self.prompt_organizacao}\n\n{parte}"
            
            # Enviar para Copilot
            resposta = self.enviar_para_copilot(texto_para_envio, f"parte {i} para organização")
            
            if resposta:
                # Extrair questões da resposta
                questões_extraidas = self.extrair_questoes_da_resposta(resposta)
                todas_questoes_organizadas.extend(questões_extraidas)
                print(f"  ✅ {len(questões_extraidas)} questões extraídas da parte {i}")
                
                # Salvar resposta intermediária para debug
                with open(f"resposta_parte_{i}.txt", "w", encoding="utf-8") as f:
                    f.write(resposta)
            else:
                print(f"  ❌ Falha ao processar parte {i}")
            
            # Limpar chat para próxima parte
            self.limpar_chat()
            
            # Pausa entre partes
            if i < len(partes_texto):
                print("  ⏳ Aguardando 10 segundos antes da próxima parte...")
                time.sleep(10)
        
        print(f"\n✅ Total de questões organizadas: {len(todas_questoes_organizadas)}")
        return todas_questoes_organizadas

    def extrair_questoes_da_resposta(self, texto_resposta):
        """Extrai questões individuais da resposta organizada do Copilot"""
        print("  🔍 Extraindo questões da resposta do Copilot...")
        
        # Padrão para identificar cada questão
        padrao_questao = r'##### QUESTÃO (\d+) #####(.*?)##### FIM QUESTÃO \1 #####'
        matches = re.findall(padrao_questao, texto_resposta, re.DOTALL)
        
        questões = []
        
        for numero, conteudo in matches:
            try:
                questao = self.parse_questao_organizada(conteudo.strip(), int(numero))
                if self.validar_questao(questao):
                    questões.append(questao)
                    print(f"    ✅ Questão {numero} extraída e validada")
                else:
                    print(f"    ⚠️ Questão {numero} não passou na validação")
            except Exception as e:
                print(f"    ❌ Erro ao processar questão {numero}: {e}")
        
        # Se não encontrou pelo padrão, tentar encontrar questões de outra forma
        if not questões:
            print("  🔍 Tentando padrão alternativo...")
            questões = self.extrair_questoes_alternativo(texto_resposta)
        
        return questões

    def extrair_questoes_alternativo(self, texto_resposta):
        """Método alternativo para extrair questões"""
        questões = []
        
        # Dividir por linhas que começam com "ENUNCIADO:"
        partes = re.split(r'(?=ENUNCIADO:)', texto_resposta)
        
        for i, parte in enumerate(partes[1:], 1):  # Pular o primeiro se for vazio
            try:
                questao = self.parse_questao_organizada(parte.strip(), i)
                if self.validar_questao(questao):
                    questões.append(questao)
            except Exception as e:
                continue
        
        return questões

    def parse_questao_organizada(self, texto_questao, numero):
        """Analisa uma questão já organizada pelo Copilot - MELHORADO"""
        questao = {
            "numero": numero,
            "texto_completo": texto_questao,
            "enunciado": "",
            "itens": {},
            "item_correto": "",
            "tem_gabarito": False,
            "explicacao": ""
        }
        
        # CORREÇÃO CRÍTICA: Limpar o enunciado para remover respostas do Copilot
        # Extrair ENUNCIADO - apenas até o início dos ITENS
        match_enunciado = re.search(r'ENUNCIADO:\s*(.*?)(?=\s*ITENS:\s*[A-E]\)|\s*ITEM CORRETO:|\s*EXPLICAÇÃO:|\s*#####)', texto_questao, re.DOTALL)
        if match_enunciado:
            enunciado_bruto = match_enunciado.group(1).strip()
            # REMOVER qualquer texto que pareça ser resposta do Copilot
            enunciado_limpo = self.limpar_enunciado(enunciado_bruto)
            questao["enunciado"] = enunciado_limpo
        else:
            # Fallback: pegar tudo após ENUNCIADO: até o primeiro item
            match_fallback = re.search(r'ENUNCIADO:\s*(.*?)(?=\s*[A-E]\)|\s*ITEM CORRETO:)', texto_questao, re.DOTALL)
            if match_fallback:
                enunciado_bruto = match_fallback.group(1).strip()
                questao["enunciado"] = self.limpar_enunciado(enunciado_bruto)
        
        # Extrair ITENS - MAIS ROBUSTO
        padrao_itens = r'([A-E])\)\s*(.*?)(?=\n[A-E]\)|\nITEM CORRETO:|\nEXPLICAÇÃO:|\n#####|$)'
        itens = re.findall(padrao_itens, texto_questao, re.DOTALL)
        for letra, texto in itens:
            # Limpar o texto do item
            texto_limpo = texto.strip()
            # Remover quebras de linha excessivas
            texto_limpo = re.sub(r'\n+', ' ', texto_limpo)
            questao["itens"][letra.upper()] = texto_limpo
        
        # Extrair ITEM CORRETO
        match_correto = re.search(r'ITEM CORRETO:\s*([A-E])', texto_questao, re.IGNORECASE)
        if match_correto:
            questao["item_correto"] = match_correto.group(1).upper()
            questao["tem_gabarito"] = True
        
        # Extrair EXPLICAÇÃO
        match_explicacao = re.search(r'EXPLICAÇÃO:\s*(.*?)(?=\n#####|$)', texto_questao, re.DOTALL)
        if match_explicacao:
            questao["explicacao"] = match_explicacao.group(1).strip()
        
        return questao

    def limpar_enunciado(self, enunciado_bruto):
        """Remove respostas do Copilot e outros textos indesejados do enunciado"""
        # Remover padrões comuns de respostas do Copilot
        padroes_remover = [
            r'Resposta fornecida pelo Copilot.*',
            r'ITEM CORRETO:.*',
            r'EXPLICAÇÃO:.*',
            r'Alternativa correta:.*',
            r'Resposta:.*',
            r'Gabarito:.*',
            r'\(Multiple-answer question\).*',
            r'\(Single-answer question\).*',
            r'- B\).*?- C\).*?- D\).*',  # Remove análises de alternativas
        ]
        
        enunciado_limpo = enunciado_bruto
        for padrao in padroes_remover:
            enunciado_limpo = re.sub(padrao, '', enunciado_limpo, flags=re.DOTALL | re.IGNORECASE)
        
        # Remover linhas em branco excessivas e espaços extras
        enunciado_limpo = re.sub(r'\n\s*\n', '\n', enunciado_limpo)
        enunciado_limpo = enunciado_limpo.strip()
        
        # Se o enunciado ainda estiver muito longo, pegar apenas o primeiro parágrafo
        if len(enunciado_limpo) > 500:
            linhas = enunciado_limpo.split('\n')
            if linhas:
                enunciado_limpo = linhas[0].strip()
        
        return enunciado_limpo

    def validar_questao(self, questao_dict):
        """Valida se uma questão tem estrutura mínima válida - MAIS RIGOROSO"""
        if not questao_dict.get('enunciado') or len(questao_dict['enunciado'].strip()) < 10:
            return False
        
        if not questao_dict.get('itens') or len(questao_dict['itens']) < 2:
            return False
        
        # Verificar se o enunciado não contém respostas do Copilot
        enunciado = questao_dict['enunciado'].lower()
        termos_proibidos = ['copilot', 'resposta fornecida', 'gabarito:', 'alternativa correta:']
        if any(termo in enunciado for termo in termos_proibidos):
            print(f"    ⚠️ Questão {questao_dict['numero']} contém resposta do Copilot no enunciado")
            return False
            
        return True

    def obter_gabaritos_com_copilot(self, todas_questoes):
        """Obtém gabaritos para questões que não têm usando Copilot - MELHORADO"""
        print("\n🎯 OBTENDO GABARITOS COM COPILOT")
        
        questões_sem_gabarito = [q for q in todas_questoes if not q['tem_gabarito']]
        
        if not questões_sem_gabarito:
            print("✅ Todas as questões já têm gabarito!")
            return todas_questoes
        
        print(f"📋 Processando {len(questões_sem_gabarito)} questões sem gabarito")
        
        # Contador regressivo
        print("⏰ Iniciando em 3 segundos...")
        for i in range(3, 0, -1):
            print(f"  {i}...")
            time.sleep(1)
        
        for i, questao in enumerate(questões_sem_gabarito, 1):
            print(f"\n{'='*40}")
            print(f"🎯 OBTENDO GABARITO QUESTÃO {questao['numero']} ({i}/{len(questões_sem_gabarito)})")
            print(f"{'='*40}")
            
            # Preparar texto da questão - APENAS enunciado e alternativas
            texto_questao = f"ENUNCIADO:\n{questao['enunciado']}\n\nALTERNATIVAS:\n"
            for letra, texto in sorted(questao['itens'].items()):
                texto_questao += f"{letra}) {texto}\n"
            
            texto_completo = f"{self.prompt_gabarito}\n\n{texto_questao}"
            
            # Enviar para Copilot
            resposta = self.enviar_para_copilot(texto_completo, f"questão {questao['numero']} para gabarito")
            
            if resposta:
                # Processar resposta do Copilot
                questao_atualizada = self.processar_resposta_gabarito(resposta, questao)
                if questao_atualizada['tem_gabarito']:
                    print(f"  ✅ Gabarito obtido: {questao_atualizada['item_correto']}")
                else:
                    print(f"  ⚠️ Não foi possível obter gabarito")
            else:
                questao_atualizada = questao
                print(f"  ❌ Falha ao obter resposta do Copilot")
            
            # Atualizar a questão na lista
            todas_questoes[questao['numero'] - 1] = questao_atualizada
            
            # Limpar chat para próxima questão
            self.limpar_chat()
            
            # Pausa entre questões
            if i < len(questões_sem_gabarito):
                print("  ⏳ Aguardando 8 segundos antes da próxima questão...")
                time.sleep(8)
        
        return todas_questoes

    def processar_resposta_gabarito(self, texto_resposta, questao_original):
        """Processa a resposta do Copilot para obter gabarito - MELHORADO"""
        questao = questao_original.copy()
        
        try:
            # Extrair ITEM CORRETO da resposta do Copilot
            padrao_correto = r'ITEM CORRETO:\s*([A-E])'
            match = re.search(padrao_correto, texto_resposta, re.IGNORECASE)
            if match:
                questao["item_correto"] = match.group(1).upper()
                questao["tem_gabarito"] = True
            
            # Extrair EXPLICAÇÃO da resposta do Copilot
            padrao_explicacao = r'EXPLICAÇÃO:\s*(.*?)(?=\s*(ITEM CORRETO:|$))'
            match_explicacao = re.search(padrao_explicacao, texto_resposta, re.DOTALL)
            if match_explicacao:
                explicacao = match_explicacao.group(1).strip()
                # Adicionar aviso que é resposta do Copilot
                if not questao_original['tem_gabarito'] and questao['tem_gabarito']:
                    explicacao = f"Resposta fornecida pelo Copilot (não oficial): {explicacao}"
                questao["explicacao"] = explicacao
            
        except Exception as e:
            print(f"  ❌ Erro ao processar resposta do gabarito: {e}")
        
        return questao

    def salvar_questoes_intermediarias(self, todas_questoes, pasta_saida):
        """Salva as questões processadas em arquivos JSON"""
        pasta_intermediaria = os.path.join(pasta_saida, "questoes_processadas")
        os.makedirs(pasta_intermediaria, exist_ok=True)
        
        for questao in todas_questoes:
            caminho_json = os.path.join(pasta_intermediaria, f"questao_{questao['numero']:03d}.json")
            with open(caminho_json, 'w', encoding='utf-8') as f:
                json.dump(questao, f, ensure_ascii=False, indent=2)
        
        print(f"💾 {len(todas_questoes)} questões salvas em {pasta_intermediaria}")

    def converter_para_formato_django(self, questoes):
        """Converte as questões para o formato Django fixtures - MELHORADO"""
        fixture_data = []
        
        question_pk = 1
        alternative_pk = 1
        correct_source_pk = 1
        
        current_time = datetime.now().strftime("%Y-%m-%dT%H:%M:%SZ")
        
        for questao in questoes:
            # Pular questões inválidas
            if not self.validar_questao_para_fixture(questao):
                print(f"  ⚠️ Pulando questão {questao['numero']} - inválida para fixture")
                continue
                
            # Adicionar objeto Question
            fixture_data.append({
                "model": "yourapp.Question",
                "pk": question_pk,
                "fields": {
                    "submitted_by": 1,
                    "reviewed_by": 2,
                    "text": questao['enunciado'],
                    "level": "HCIA",
                    "has_answer": questao['tem_gabarito'],
                    "has_multiple_answers": False,
                    "track": "Computing",
                    "weight": "1.00",
                    "approved_at": current_time if questao['tem_gabarito'] else None,
                    "last_update": current_time
                }
            })
            
            # Adicionar objetos Alternative
            for letra, texto in sorted(questao['itens'].items()):
                is_correct = (letra == questao['item_correto'])
                
                fixture_data.append({
                    "model": "yourapp.Alternative",
                    "pk": alternative_pk,
                    "fields": {
                        "question": question_pk,
                        "text": texto,
                        "is_correct": is_correct
                    }
                })
                
                # Se for a alternativa correta, adicionar CorrectAnswersSources
                if is_correct and questao['tem_gabarito']:
                    source_text = "Documento oficial"
                    if "Resposta fornecida pelo Copilot" in questao.get('explicacao', ''):
                        source_text = "Resposta fornecida pelo Copilot (não oficial)"
                    
                    fixture_data.append({
                        "model": "yourapp.CorrectAnswersSources",
                        "pk": correct_source_pk,
                        "fields": {
                            "alternative": alternative_pk,
                            "source": source_text
                        }
                    })
                    correct_source_pk += 1
                
                alternative_pk += 1
            
            question_pk += 1
        
        return fixture_data

    def validar_questao_para_fixture(self, questao):
        """Validação mais rigorosa para inclusão no fixture"""
        if not questao.get('enunciado') or len(questao['enunciado'].strip()) < 10:
            return False
        
        if not questao.get('itens') or len(questao['itens']) < 2:
            return False
        
        # Verificar se tem gabarito
        if not questao.get('tem_gabarito') or not questao.get('item_correto'):
            return False
            
        # Verificar se o item correto existe nas alternativas
        if questao['item_correto'] not in questao['itens']:
            return False
            
        return True
    
    def salvar_json_django_fixture(self, fixture_data, pasta_saida):
        """Salva as questões no formato Django fixture"""
        os.makedirs(pasta_saida, exist_ok=True)
        
        caminho_json = os.path.join(pasta_saida, "questions_fixture.json")
        
        with open(caminho_json, 'w', encoding='utf-8') as f:
            json.dump(fixture_data, f, ensure_ascii=False, indent=2)
        
        print(f"📄 Arquivo Django fixture salvo em: {caminho_json}")
        print(f"📊 Total de objetos salvos: {len(fixture_data)}")
        
        # Estatísticas
        questions_count = sum(1 for item in fixture_data if item['model'] == 'yourapp.Question')
        alternatives_count = sum(1 for item in fixture_data if item['model'] == 'yourapp.Alternative')
        sources_count = sum(1 for item in fixture_data if item['model'] == 'yourapp.CorrectAnswersSources')
        
        print(f"📈 Estatísticas:")
        print(f"   - Questions: {questions_count}")
        print(f"   - Alternatives: {alternatives_count}")
        print(f"   - CorrectAnswersSources: {sources_count}")
        
        return caminho_json
    
    def gerar_relatorio_processamento(self, todas_questoes, fixture_data, pasta_saida):
        """Gera um relatório detalhado do processamento"""
        caminho_relatorio = os.path.join(pasta_saida, "RELATORIO_PROCESSAMENTO.txt")
        
        questões_com_gabarito = sum(1 for q in todas_questoes if q['tem_gabarito'])
        questões_sem_gabarito = sum(1 for q in todas_questoes if not q['tem_gabarito'])
        questões_validas_fixture = sum(1 for q in todas_questoes if self.validar_questao_para_fixture(q))
        
        with open(caminho_relatorio, 'w', encoding='utf-8') as f:
            f.write("RELATÓRIO DE PROCESSAMENTO DE QUESTÕES (TEXTO DOCX)\n")
            f.write("=" * 60 + "\n\n")
            
            f.write(f"Total de questões identificadas: {len(todas_questoes)}\n")
            f.write(f"Questões com gabarito no documento: {questões_com_gabarito}\n")
            f.write(f"Questões sem gabarito (processadas pelo Copilot): {questões_sem_gabarito}\n")
            f.write(f"Questões válidas para fixture: {questões_validas_fixture}\n")
            f.write(f"Total de objetos no fixture: {len(fixture_data)}\n\n")
            
            f.write("DETALHES DAS QUESTÕES:\n")
            f.write("-" * 30 + "\n")
            
            for questao in todas_questoes:
                f.write(f"\nQuestão {questao['numero']}:\n")
                f.write(f"  Enunciado: {len(questao['enunciado'])} caracteres\n")
                f.write(f"  Alternativas: {len(questao['itens'])}\n")
                f.write(f"  Tem gabarito: {'Sim' if questao['tem_gabarito'] else 'Não (Copilot)'}\n")
                if questao['tem_gabarito']:
                    f.write(f"  Item correto: {questao['item_correto']}\n")
                if questao.get('explicacao'):
                    f.write(f"  Explicação: {len(questao['explicacao'])} caracteres\n")
                f.write(f"  Válida para fixture: {'Sim' if self.validar_questao_para_fixture(questao) else 'Não'}\n")
        
        print(f"📊 Relatório de processamento salvo em: {caminho_relatorio}")
    
    def processar_documento(self, arquivo_docx, pasta_saida="questoes_processadas_docx"):
        """Processa o documento DOCX usando abordagem em duas fases - MELHORADO"""
        print("🚀 INICIANDO PROCESSAMENTO INTELIGENTE DO DOCX")
        print("=" * 50)
        print("📝 ABORDAGEM: ORGANIZAÇÃO + GABARITO (CORRIGIDA)")
        print("=" * 50)
        
        if not os.path.exists(arquivo_docx):
            print(f"❌ Arquivo não encontrado: {arquivo_docx}")
            return
        
        if not self.verificar_dependencias():
            return
        
        # 1. Extrair texto do DOCX
        texto_completo = self.extrair_texto_docx(arquivo_docx)
        if not texto_completo:
            return
        
        # 2. FASE 1: Organizar questões com Copilot
        print("\n" + "="*60)
        print("🎯 FASE 1: ORGANIZANDO QUESTÕES COM COPILOT")
        print("="*60)
        todas_questoes = self.organizar_questoes_com_copilot(texto_completo)
        
        if not todas_questoes:
            print("❌ Nenhuma questão foi organizada")
            return
        
        # 3. FASE 2: Obter gabaritos com Copilot
        print("\n" + "="*60)
        print("🎯 FASE 2: OBTENDO GABARITOS COM COPILOT")
        print("="*60)
        todas_questoes_com_gabarito = self.obter_gabaritos_com_copilot(todas_questoes)
        
        # 4. Salvar questões intermediárias
        self.salvar_questoes_intermediarias(todas_questoes_com_gabarito, pasta_saida)
        
        # 5. Converter para formato Django
        print(f"\n🔄 CONVERTENDO PARA FORMATO DJANGO FIXTURE...")
        fixture_data = self.converter_para_formato_django(todas_questoes_com_gabarito)
        
        # 6. Salvar resultados
        print(f"\n💾 SALVANDO RESULTADOS...")
        caminho_json = self.salvar_json_django_fixture(fixture_data, pasta_saida)
        self.gerar_relatorio_processamento(todas_questoes_com_gabarito, fixture_data, pasta_saida)
        
        # 7. Relatório final
        print(f"\n✅ PROCESSAMENTO CONCLUÍDO!")
        print(f"📄 Total de questões extraídas: {len(todas_questoes_com_gabarito)}")
        print(f"📄 Questões válidas no fixture: {sum(1 for q in todas_questoes_com_gabarito if self.validar_questao_para_fixture(q))}")
        print(f"📁 Pasta de saída: {pasta_saida}")
        print(f"📝 Arquivo principal: {os.path.basename(caminho_json)}")

def main():
    """Função principal"""
    extrator = ExtratorQuestoesDOCX()
    
    # Configurações
    arquivo_docx = "ICT Computing - Questions Database.docx"
    pasta_saida = "questoes_processadas_docx"
    
    # Executar processamento
    extrator.processar_documento(arquivo_docx, pasta_saida)

if __name__ == "__main__":
    main()
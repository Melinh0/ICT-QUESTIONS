#!/usr/bin/env python3
# -*- coding: utf-8 -*-

import zipfile
import os
import re
import json
import time
import pyautogui
import pyperclip
import glob
from datetime import datetime

class ExtratorTextoComCopilot:
    def __init__(self):
        self.coordenadas = {
            'anexar_arquivo': (3215, 607),
            'adicionar_imagem': (3300, 658),
            'pasta_pessoal': (2952, 455),
            'pesquisar_pasta': (3710, 368),
            'clicar_digitar': (3304, 420),
            'pasta_tutorbots': (3127, 482),
            'pasta_imagens': (3139, 526),
            'pasta_questoes': (3139, 528),
            'icone_pesquisa': (3708, 375),
            'pesquisar_imagem': (3299, 415),
            'selecionar_imagem': (3219, 485),
            'enviar_prompt': (3229, 629),
            'enviar_mensagem': (3788, 792),
            'descer_resposta': (3501, 901),
            'copiar_resposta': (3319, 892),
            'icone_chat': (3149, 441),
            'excluir_chat': (3172, 518),
            'confirmar_exclusao': (3358, 625)
        }
        self.prompt_copilot = """Preciso que transcreva TODO o texto da imagem da questão de forma COMPLETA e organize EXATAMENTE neste formato:

ENUNCIADO: [todo o texto do enunciado aqui]

ITENS:
A) [texto completo do item A]
B) [texto completo do item B]
C) [texto completo do item C]
D) [texto completo do item D]
E) [texto completo do item E] (se houver)

ITEM CORRETO: [letra do item correto]

EXPLICAÇÃO: [texto completo da explicação se houver na imagem]

Seja DIRETO e copie APENAS a transcrição fiel do texto da imagem. Não interprete, não explique, apenas transcreva. Se tiver questões que na verdade são Verdadeiro ou Falso, os itens da questão são esses."""
        
    def verificar_dependencias(self):
        """Verifica se todas as dependências estão disponíveis"""
        try:
            import pyautogui
            import pyperclip
            print("✅ PyAutoGUI e PyPerClip encontrados")
            return True
        except ImportError as e:
            print(f"❌ Dependências não encontradas: {e}")
            print("Instale com: pip install pyautogui pyperclip")
            return False
    
    def extrair_imagens_docx(self, arquivo_docx, pasta_destino):
        """Extrai todas as imagens do documento DOCX"""
        print(f"🖼️ Extraindo imagens de: {arquivo_docx}")
        
        os.makedirs(pasta_destino, exist_ok=True)
        imagens_extraidas = []
        
        try:
            with zipfile.ZipFile(arquivo_docx, 'r') as docx_zip:
                arquivos = docx_zip.namelist()
                imagens = [f for f in arquivos if f.startswith('word/media/') and os.path.basename(f)]
                
                print(f"📸 Encontradas {len(imagens)} imagens no documento")
                
                for i, caminho_imagem in enumerate(imagens, 1):
                    try:
                        dados_imagem = docx_zip.read(caminho_imagem)
                        _, extensao = os.path.splitext(caminho_imagem)
                        if not extensao:
                            extensao = '.png'
                        
                        nome_arquivo = f'questao_{i:03d}{extensao}'
                        caminho_completo = os.path.join(pasta_destino, nome_arquivo)
                        
                        with open(caminho_completo, 'wb') as f:
                            f.write(dados_imagem)
                        
                        imagens_extraidas.append({
                            'caminho': caminho_completo,
                            'nome': nome_arquivo,
                            'numero': i
                        })
                        print(f"  ✅ {nome_arquivo}")
                        
                    except Exception as e:
                        print(f"  ⚠️ Erro na imagem {i}: {e}")
                        
        except Exception as e:
            print(f"❌ Erro ao extrair imagens do DOCX: {e}")
            
        return imagens_extraidas
    
    def verificar_checkpoint(self, pasta_saida):
        """Verifica quais questões já foram processadas e retorna o último ponto"""
        pasta_intermediaria = os.path.join(pasta_saida, "respostas_intermediarias")
        
        if not os.path.exists(pasta_intermediaria):
            return 0, []
        
        # Encontrar todos os arquivos de resposta
        arquivos_resposta = glob.glob(os.path.join(pasta_intermediaria, "resposta_questao_*.txt"))
        
        # Extrair números das questões já processadas
        numeros_processados = []
        for arquivo in arquivos_resposta:
            try:
                numero = int(re.search(r'resposta_questao_(\d+)\.txt', arquivo).group(1))
                numeros_processados.append(numero)
            except:
                continue
        
        if numeros_processados:
            ultimo_numero = max(numeros_processados)
            print(f"📌 Checkpoint encontrado: {len(numeros_processados)} questões já processadas")
            print(f"📌 Última questão processada: {ultimo_numero}")
            return ultimo_numero, numeros_processados
        else:
            return 0, []
    
    def carregar_questoes_existentes(self, pasta_saida, numeros_processados):
        """Carrega as questões já processadas a partir dos arquivos intermediários"""
        if not numeros_processados:
            return []
        
        pasta_intermediaria = os.path.join(pasta_saida, "respostas_intermediarias")
        questões_existentes = []
        
        for numero in numeros_processados:
            caminho_arquivo = os.path.join(pasta_intermediaria, f"resposta_questao_{numero:03d}.txt")
            if os.path.exists(caminho_arquivo):
                try:
                    with open(caminho_arquivo, 'r', encoding='utf-8') as f:
                        texto_resposta = f.read()
                    
                    # Extrair informações da resposta
                    questao = self.extrair_informacoes_resposta(texto_resposta, numero)
                    if questao:
                        questões_existentes.append(questao)
                        print(f"  ✅ Questão {numero} carregada do checkpoint")
                    
                except Exception as e:
                    print(f"  ⚠️ Erro ao carregar questão {numero}: {e}")
        
        return questões_existentes
    
    def clicar_coordenada(self, coordenada, descricao="", duplo_clique=False):
        """Clica em uma coordenada específica na tela"""
        print(f"  🖱️ Clicando em {descricao}: {coordenada}")
        x, y = coordenada
        
        # Mover o mouse para a coordenada
        pyautogui.moveTo(x, y, duration=0.5)
        time.sleep(0.5)
        
        if duplo_clique:
            pyautogui.doubleClick()
        else:
            pyautogui.click()
        
        time.sleep(1)
    
    def digitar_texto(self, texto, descricao=""):
        """Digita texto no campo atual"""
        print(f"  ⌨️ Digitando {descricao}: {texto}")
        pyautogui.write(texto, interval=0.05)
        time.sleep(1)
    
    def colar_texto(self, texto, descricao=""):
        """Copia texto para a área de transferência e cola"""
        print(f"  📋 Colando {descricao} ({len(texto)} caracteres)")
        pyperclip.copy(texto)
        time.sleep(0.5)
        pyautogui.hotkey('ctrl', 'v')
        time.sleep(1)
    
    def salvar_resposta_intermediaria(self, texto, numero_questao, pasta_saida):
        """Salva a resposta do Copilot em um arquivo TXT intermediário"""
        pasta_intermediaria = os.path.join(pasta_saida, "respostas_intermediarias")
        os.makedirs(pasta_intermediaria, exist_ok=True)
        
        caminho_txt = os.path.join(pasta_intermediaria, f"resposta_questao_{numero_questao:03d}.txt")
        
        with open(caminho_txt, 'w', encoding='utf-8') as f:
            f.write(f"RESPOSTA DA QUESTÃO {numero_questao}\n")
            f.write("=" * 50 + "\n\n")
            f.write(texto)
        
        print(f"  💾 Resposta intermediária salva: {os.path.basename(caminho_txt)}")
        return caminho_txt
    
    def processar_imagem_copilot(self, imagem_info, pasta_saida):
        """Processa uma imagem usando o Copilot via automação de interface"""
        print(f"\n🔄 PROCESSANDO IMAGEM {imagem_info['numero']}: {imagem_info['nome']}")
        
        try:
            # 1. Clicar para anexar arquivo
            self.clicar_coordenada(
                self.coordenadas['anexar_arquivo'], 
                "Anexar arquivo"
            )
            
            # 2. Clicar em "Adicionar imagem"
            self.clicar_coordenada(
                self.coordenadas['adicionar_imagem'], 
                "Adicionar imagem"
            )
            
            # 3. Clicar na pasta pessoal
            self.clicar_coordenada(
                self.coordenadas['pasta_pessoal'], 
                "Pasta pessoal"
            )
            
            # 4. Clicar para pesquisar pasta
            self.clicar_coordenada(
                self.coordenadas['pesquisar_pasta'], 
                "Pesquisar pasta"
            )
            
            # 5. Clicar para digitar
            self.clicar_coordenada(
                self.coordenadas['clicar_digitar'], 
                "Campo de digitação"
            )
            
            # 6. Digitar "Tutorbots"
            self.digitar_texto("Tutorbots", "nome da pasta Tutorbots")
            
            # 7. Clicar duas vezes na pasta Tutorbots
            self.clicar_coordenada(
                self.coordenadas['pasta_tutorbots'], 
                "Pasta Tutorbots", 
                duplo_clique=True
            )
            
            # 8. Clicar duas vezes na pasta de imagens
            self.clicar_coordenada(
                self.coordenadas['pasta_imagens'], 
                "Pasta de imagens", 
                duplo_clique=True
            )
            
            # 9. Clicar duas vezes na pasta de questões
            self.clicar_coordenada(
                self.coordenadas['pasta_questoes'], 
                "Pasta de questões", 
                duplo_clique=True
            )
            
            # 10. Clicar no ícone de pesquisa
            self.clicar_coordenada(
                self.coordenadas['icone_pesquisa'], 
                "Ícone de pesquisa"
            )
            
            # 11. Clicar na aba de pesquisar imagem
            self.clicar_coordenada(
                self.coordenadas['pesquisar_imagem'], 
                "Campo pesquisar imagem"
            )
            
            # 12. Digitar o nome da imagem
            nome_imagem = imagem_info['nome'].replace('.png', '').replace('.jpg', '')
            self.digitar_texto(nome_imagem, f"nome da imagem {nome_imagem}")
            
            # 13. Clicar duas vezes na imagem para anexar
            self.clicar_coordenada(
                self.coordenadas['selecionar_imagem'], 
                "Selecionar imagem", 
                duplo_clique=True
            )
            
            time.sleep(3)  # Aguardar upload da imagem
            
            # 14. Clicar no campo de prompt
            self.clicar_coordenada(
                self.coordenadas['enviar_prompt'], 
                "Campo de prompt"
            )
            
            # 15. COLAR o prompt (em vez de digitar)
            self.colar_texto(self.prompt_copilot, "prompt para transcrição")
            
            # 16. Enviar mensagem (coordenada corrigida)
            self.clicar_coordenada(
                self.coordenadas['enviar_mensagem'], 
                "Enviar mensagem"
            )
            
            print("  ⏳ Aguardando resposta do Copilot (30 segundos)...")
            time.sleep(15)  # Aguardar resposta
            
            # 17. Descer a resposta
            self.clicar_coordenada(
                self.coordenadas['descer_resposta'], 
                "Descer resposta"
            )
            time.sleep(2)
            
            # 18. Copiar resposta
            self.clicar_coordenada(
                self.coordenadas['copiar_resposta'], 
                "Copiar resposta"
            )
            time.sleep(2)
            
            # Obter texto da área de transferência usando pyperclip
            try:
                texto_copiado = pyperclip.paste()
                print(f"  ✅ Texto copiado ({len(texto_copiado)} caracteres)")
                
                # Salvar resposta intermediária
                caminho_intermediario = self.salvar_resposta_intermediaria(
                    texto_copiado, 
                    imagem_info['numero'], 
                    pasta_saida
                )
                
            except Exception as e:
                print(f"  ❌ Erro ao acessar área de transferência: {e}")
                # Tentar método alternativo
                texto_copiado = self.tentar_copiar_texto_alternativo()
            
            # 19. Limpar chat para próxima imagem
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
            
            return self.extrair_informacoes_resposta(texto_copiado, imagem_info['numero'])
            
        except Exception as e:
            print(f"❌ Erro ao processar imagem {imagem_info['nome']}: {e}")
            return None
    
    def tentar_copiar_texto_alternativo(self):
        """Método alternativo para copiar texto se pyperclip falhar"""
        print("  🔄 Tentando método alternativo de cópia...")
        try:
            # Selecionar tudo e copiar
            pyautogui.hotkey('ctrl', 'a')
            time.sleep(1)
            pyautogui.hotkey('ctrl', 'c')
            time.sleep(1)
            return pyperclip.paste()
        except:
            return "Texto não pôde ser copiado"
    
    def extrair_informacoes_resposta(self, texto_resposta, numero_questao):
        """Extrai as informações da resposta do Copilot e organiza em estrutura"""
        print(f"  🔍 Processando resposta da questão {numero_questao}")
        
        questao = {
            "numero": numero_questao,
            "enunciado": "",
            "itens": {},
            "item_correto": "",
            "explicacao": ""
        }
        
        try:
            # Extrair ENUNCIADO
            inicio_enunciado = texto_resposta.find("ENUNCIADO:")
            if inicio_enunciado != -1:
                inicio_enunciado += len("ENUNCIADO:")
                fim_enunciado = texto_resposta.find("ITENS:")
                if fim_enunciado == -1:
                    fim_enunciado = texto_resposta.find("ITEM CORRETO:")
                if fim_enunciado == -1:
                    fim_enunciado = len(texto_resposta)
                
                questao["enunciado"] = texto_resposta[inicio_enunciado:fim_enunciado].strip()
            
            # Extrair ITENS
            inicio_itens = texto_resposta.find("ITENS:")
            if inicio_itens != -1:
                inicio_itens += len("ITENS:")
                fim_itens = texto_resposta.find("ITEM CORRETO:")
                if fim_itens == -1:
                    fim_itens = texto_resposta.find("EXPLICAÇÃO:")
                if fim_itens == -1:
                    fim_itens = len(texto_resposta)
                
                texto_itens = texto_resposta[inicio_itens:fim_itens].strip()
                
                # Extrair itens individuais (A, B, C, D, E)
                for letra in ['A', 'B', 'C', 'D', 'E']:
                    padrao = f"{letra}\)\s*(.*?)(?=\n[A-E]\)|$)"
                    matches = re.findall(padrao, texto_itens, re.DOTALL)
                    if matches:
                        questao["itens"][letra] = matches[0].strip()
            
            # Extrair ITEM CORRETO
            inicio_correto = texto_resposta.find("ITEM CORRETO:")
            if inicio_correto != -1:
                inicio_correto += len("ITEM CORRETO:")
                fim_correto = texto_resposta.find("EXPLICAÇÃO:")
                if fim_correto == -1:
                    fim_correto = len(texto_resposta)
                
                item_correto = texto_resposta[inicio_correto:fim_correto].strip()
                # Extrair apenas a letra (A, B, C, D, E)
                match_letra = re.search(r'([A-E])', item_correto)
                if match_letra:
                    questao["item_correto"] = match_letra.group(1)
            
            # Extrair EXPLICAÇÃO
            inicio_explicacao = texto_resposta.find("EXPLICAÇÃO:")
            if inicio_explicacao != -1:
                inicio_explicacao += len("EXPLICAÇÃO:")
                questao["explicacao"] = texto_resposta[inicio_explicacao:].strip()
            
            print(f"  ✅ Questão {numero_questao} processada:")
            print(f"     - Enunciado: {len(questao['enunciado'])} caracteres")
            print(f"     - Itens: {len(questao['itens'])} encontrados")
            print(f"     - Item correto: {questao['item_correto']}")
            print(f"     - Explicação: {len(questao['explicacao'])} caracteres")
            
            return questao
            
        except Exception as e:
            print(f"  ❌ Erro ao extrair informações da resposta: {e}")
            return questao
    
    def processar_imagens_restantes(self, imagens, pasta_saida, ultimo_numero_processado):
        """Processa apenas as imagens que ainda não foram processadas"""
        print(f"\n🚀 CONTINUANDO PROCESSAMENTO A PARTIR DA QUESTÃO {ultimo_numero_processado + 1}")
        print("=" * 60)
        
        # Filtrar apenas imagens não processadas
        imagens_restantes = [img for img in imagens if img['numero'] > ultimo_numero_processado]
        
        if not imagens_restantes:
            print("✅ Todas as imagens já foram processadas!")
            return []
        
        print(f"📋 Processando {len(imagens_restantes)} imagens restantes")
        
        todas_questoes = []
        
        # Contador regressivo antes de começar
        print("⏰ Iniciando em 5 segundos... Posicione o mouse!")
        for i in range(5, 0, -1):
            print(f"  {i}...")
            time.sleep(1)
        
        for imagem_info in imagens_restantes:
            questao = self.processar_imagem_copilot(imagem_info, pasta_saida)
            if questao:
                todas_questoes.append(questao)
            
            # Pequena pausa entre processamentos
            time.sleep(2)
        
        return todas_questoes
    
    def converter_para_formato_django(self, questoes):
        """Converte as questões para o formato Django fixtures"""
        fixture_data = []
        
        # Contadores para PKs
        question_pk = 1
        alternative_pk = 1
        correct_source_pk = 1
        
        current_time = datetime.now().strftime("%Y-%m-%dT%H:%M:%SZ")
        
        for questao in questoes:
            # Adicionar objeto Question
            fixture_data.append({
                "model": "yourapp.Question",
                "pk": question_pk,
                "fields": {
                    "submitted_by": 1,
                    "reviewed_by": 2,
                    "text": questao['enunciado'],
                    "level": "HCIA",
                    "has_answer": True,
                    "has_multiple_answers": False,
                    "track": "Cloud",
                    "weight": "1.00",
                    "approved_at": current_time,
                    "last_update": current_time
                }
            })
            
            # Adicionar objetos Alternative
            for letra, texto in questao['itens'].items():
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
                if is_correct:
                    fixture_data.append({
                        "model": "yourapp.CorrectAnswersSources",
                        "pk": correct_source_pk,
                        "fields": {
                            "alternative": alternative_pk,
                            "source": "https://support.huaweicloud.com/hcs-introduction/index.html"
                        }
                    })
                    correct_source_pk += 1
                
                alternative_pk += 1
            
            question_pk += 1
        
        return fixture_data
    
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
    
    def gerar_relatorio_processamento(self, fixture_data, pasta_saida):
        """Gera um relatório detalhado do processamento"""
        caminho_relatorio = os.path.join(pasta_saida, "RELATORIO_PROCESSAMENTO.txt")
        
        questions_count = sum(1 for item in fixture_data if item['model'] == 'yourapp.Question')
        alternatives_count = sum(1 for item in fixture_data if item['model'] == 'yourapp.Alternative')
        sources_count = sum(1 for item in fixture_data if item['model'] == 'yourapp.CorrectAnswersSources')
        
        with open(caminho_relatorio, 'w', encoding='utf-8') as f:
            f.write("RELATÓRIO DE PROCESSAMENTO DE QUESTÕES (FORMATO DJANGO)\n")
            f.write("=" * 60 + "\n\n")
            
            f.write(f"Total de questões processadas: {questions_count}\n")
            f.write(f"Total de alternativas: {alternatives_count}\n")
            f.write(f"Total de fontes de resposta correta: {sources_count}\n")
            f.write(f"Total de objetos no fixture: {len(fixture_data)}\n\n")
            
            f.write("ESTRUTURA DO FIXTURE:\n")
            f.write("-" * 30 + "\n")
            
            for i in range(1, questions_count + 1):
                question_alternatives = [alt for alt in fixture_data 
                                       if alt['model'] == 'yourapp.Alternative' and alt['fields']['question'] == i]
                correct_alternative = [alt for alt in question_alternatives if alt['fields']['is_correct']]
                
                f.write(f"\nQuestão {i}:\n")
                f.write(f"  Alternativas: {len(question_alternatives)}\n")
                f.write(f"  Correta: {'Sim' if correct_alternative else 'Não'}\n")
                if correct_alternative:
                    f.write(f"  ID da alternativa correta: {correct_alternative[0]['pk']}\n")
        
        print(f"📊 Relatório de processamento salvo em: {caminho_relatorio}")
    
    def processar_documento(self, arquivo_docx, pasta_saida="questoes_processadas"):
        """Processa o documento completo usando Copilot com sistema de checkpoint"""
        print("🚀 INICIANDO EXTRAÇÃO DE QUESTÕES COM COPILOT")
        print("=" * 50)
        
        if not os.path.exists(arquivo_docx):
            print(f"❌ Arquivo não encontrado: {arquivo_docx}")
            return
        
        if not self.verificar_dependencias():
            return
        
        # 1. Verificar checkpoint
        ultimo_numero, numeros_processados = self.verificar_checkpoint(pasta_saida)
        
        # 2. Extrair imagens do DOCX
        pasta_imagens = os.path.join(pasta_saida, "imagens_extraidas")
        imagens = self.extrair_imagens_docx(arquivo_docx, pasta_imagens)
        
        if not imagens:
            print("❌ Nenhuma imagem encontrada no documento")
            return
        
        # 3. Carregar questões já processadas
        if ultimo_numero > 0:
            print(f"📥 Carregando questões já processadas...")
            questoes_existentes = self.carregar_questoes_existentes(pasta_saida, numeros_processados)
        else:
            questoes_existentes = []
        
        # 4. Processar imagens restantes
        novas_questoes = self.processar_imagens_restantes(imagens, pasta_saida, ultimo_numero)
        
        # 5. Combinar questões existentes com novas
        todas_questoes = questoes_existentes + novas_questoes
        
        if not todas_questoes:
            print("❌ Nenhuma questão foi processada")
            return
        
        # 6. Converter para formato Django fixture
        print(f"\n🔄 CONVERTENDO PARA FORMATO DJANGO FIXTURE...")
        fixture_data = self.converter_para_formato_django(todas_questoes)
        
        # 7. Salvar resultados
        print(f"\n💾 SALVANDO RESULTADOS...")
        
        caminho_json = self.salvar_json_django_fixture(fixture_data, pasta_saida)
        self.gerar_relatorio_processamento(fixture_data, pasta_saida)
        
        # 8. Relatório final
        print(f"\n✅ PROCESSAMENTO CONCLUÍDO!")
        print(f"📊 Total de imagens no documento: {len(imagens)}")
        print(f"📄 Questões processadas anteriormente: {len(questoes_existentes)}")
        print(f"📄 Novas questões processadas: {len(novas_questoes)}")
        print(f"📄 Total de questões extraídas: {len(todas_questoes)}")
        print(f"📁 Pasta de saída: {pasta_saida}")
        print(f"📝 Arquivo principal: {os.path.basename(caminho_json)}")

def main():
    """Função principal"""
    extrator = ExtratorTextoComCopilot()
    
    # Configurações
    arquivo_docx = "ICT Computing - Questions Database.docx"
    pasta_saida = "questoes_processadas_copilot"
    
    # Executar processamento
    extrator.processar_documento(arquivo_docx, pasta_saida)

if __name__ == "__main__":
    main()
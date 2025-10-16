#!/usr/bin/env python3
# -*- coding: utf-8 -*-

import json
import os
from datetime import datetime

class UnificadorQuestoes:
    def __init__(self):
        self.questoes_unificadas = []
        
    def carregar_json(self, caminho_arquivo):
        """Carrega um arquivo JSON"""
        try:
            with open(caminho_arquivo, 'r', encoding='utf-8') as f:
                return json.load(f)
        except Exception as e:
            print(f"❌ Erro ao carregar {caminho_arquivo}: {e}")
            return None
    
    def extrair_questoes_do_fixture(self, fixture_data):
        """Extrai questões individuais do formato Django fixture"""
        questoes = {}
        
        # Primeiro, coletar todas as questões
        for item in fixture_data:
            if item['model'] == 'yourapp.Question':
                questao_id = item['pk']
                questoes[questao_id] = {
                    'id': questao_id,
                    'enunciado': item['fields']['text'],
                    'alternativas': {},
                    'item_correto': None,
                    'fonte': None,
                    'has_answer': item['fields']['has_answer'],
                    'approved_at': item['fields']['approved_at']
                }
        
        # Depois, coletar alternativas
        for item in fixture_data:
            if item['model'] == 'yourapp.Alternative':
                questao_id = item['fields']['question']
                alternativa_id = item['pk']
                letra = self.obter_letra_alternativa(alternativa_id, fixture_data)
                
                if questao_id in questoes:
                    questoes[questao_id]['alternativas'][letra] = {
                        'texto': item['fields']['text'],
                        'is_correct': item['fields']['is_correct']
                    }
                    
                    # Se for a correta, marcar como item correto
                    if item['fields']['is_correct']:
                        questoes[questao_id]['item_correto'] = letra
        
        # Coletar fontes das respostas
        for item in fixture_data:
            if item['model'] == 'yourapp.CorrectAnswersSources':
                alternativa_id = item['fields']['alternative']
                # Encontrar a questão correspondente
                for questao in questoes.values():
                    if alternativa_id in [alt_id for alt_id in self.obter_ids_alternativas(questao['id'], fixture_data)]:
                        questao['fonte'] = item['fields']['source']
                        break
        
        return list(questoes.values())
    
    def obter_letra_alternativa(self, alternativa_id, fixture_data):
        """Obtém a letra da alternativa baseado na ordem"""
        # Encontrar todas as alternativas da mesma questão
        questao_id = None
        for item in fixture_data:
            if item['model'] == 'yourapp.Alternative' and item['pk'] == alternativa_id:
                questao_id = item['fields']['question']
                break
        
        if not questao_id:
            return 'A'
        
        # Coletar todas as alternativas da questão e ordenar por PK
        alternativas_questao = []
        for item in fixture_data:
            if item['model'] == 'yourapp.Alternative' and item['fields']['question'] == questao_id:
                alternativas_questao.append((item['pk'], item['fields']['text']))
        
        alternativas_questao.sort(key=lambda x: x[0])
        
        # Mapear para letras
        letras = ['A', 'B', 'C', 'D', 'E']
        for i, (alt_id, texto) in enumerate(alternativas_questao):
            if alt_id == alternativa_id:
                return letras[i] if i < len(letras) else 'A'
        
        return 'A'
    
    def obter_ids_alternativas(self, questao_id, fixture_data):
        """Obtém todos os IDs de alternativas de uma questão"""
        ids = []
        for item in fixture_data:
            if item['model'] == 'yourapp.Alternative' and item['fields']['question'] == questao_id:
                ids.append(item['pk'])
        return sorted(ids)
    
    def unificar_questoes(self, questoes1, questoes2):
        """Unifica duas listas de questões, evitando duplicatas"""
        todas_questoes = []
        enunciados_vistos = set()
        
        # Adicionar questões do primeiro arquivo
        for questao in questoes1:
            enunciado_hash = self.hash_enunciado(questao['enunciado'])
            if enunciado_hash not in enunciados_vistos:
                todas_questoes.append(questao)
                enunciados_vistos.add(enunciado_hash)
        
        # Adicionar questões do segundo arquivo (apenas se não forem duplicatas)
        for questao in questoes2:
            enunciado_hash = self.hash_enunciado(questao['enunciado'])
            if enunciado_hash not in enunciados_vistos:
                todas_questoes.append(questao)
                enunciados_vistos.add(enunciado_hash)
        
        return todas_questoes
    
    def hash_enunciado(self, enunciado):
        """Cria um hash simplificado do enunciado para detectar duplicatas"""
        # Remover espaços extras e converter para minúsculas
        texto_limpo = ' '.join(enunciado.strip().lower().split())
        # Manter apenas os primeiros 100 caracteres para o hash
        return texto_limpo[:100]
    
    def converter_para_formato_django(self, todas_questoes):
        """Converte as questões unificadas para o formato Django fixture"""
        fixture_data = []
        
        question_pk = 1
        alternative_pk = 1
        correct_source_pk = 1
        
        current_time = datetime.now().strftime("%Y-%m-%dT%H:%M:%SZ")
        
        for questao in todas_questoes:
            # Adicionar objeto Question
            fixture_data.append({
                "model": "yourapp.Question",
                "pk": question_pk,
                "fields": {
                    "submitted_by": 1,
                    "reviewed_by": 2,
                    "text": questao['enunciado'],
                    "level": "HCIA",
                    "has_answer": questao['has_answer'],
                    "has_multiple_answers": False,
                    "track": "Computing",
                    "weight": "1.00",
                    "approved_at": questao.get('approved_at', current_time) if questao['has_answer'] else None,
                    "last_update": current_time
                }
            })
            
            # Adicionar objetos Alternative
            for letra, alt_info in sorted(questao['alternativas'].items()):
                is_correct = (letra == questao['item_correto'])
                
                fixture_data.append({
                    "model": "yourapp.Alternative",
                    "pk": alternative_pk,
                    "fields": {
                        "question": question_pk,
                        "text": alt_info['texto'],
                        "is_correct": is_correct
                    }
                })
                
                # Se for a alternativa correta, adicionar CorrectAnswersSources
                if is_correct and questao['has_answer'] and questao.get('fonte'):
                    fixture_data.append({
                        "model": "yourapp.CorrectAnswersSources",
                        "pk": correct_source_pk,
                        "fields": {
                            "alternative": alternative_pk,
                            "source": questao['fonte']
                        }
                    })
                    correct_source_pk += 1
                
                alternative_pk += 1
            
            question_pk += 1
        
        return fixture_data
    
    def gerar_relatorio_unificacao(self, questoes1, questoes2, questoes_unificadas, pasta_saida):
        """Gera um relatório detalhado da unificação"""
        caminho_relatorio = os.path.join(pasta_saida, "RELATORIO_UNIFICACAO.txt")
        
        with open(caminho_relatorio, 'w', encoding='utf-8') as f:
            f.write("RELATÓRIO DE UNIFICAÇÃO DE QUESTÕES\n")
            f.write("=" * 50 + "\n\n")
            
            f.write(f"Questões do primeiro arquivo (imagens): {len(questoes1)}\n")
            f.write(f"Questões do segundo arquivo (DOCX): {len(questoes2)}\n")
            f.write(f"Questões unificadas (sem duplicatas): {len(questoes_unificadas)}\n")
            f.write(f"Duplicatas removidas: {len(questoes1) + len(questoes2) - len(questoes_unificadas)}\n\n")
            
            f.write("ESTATÍSTICAS DETALHADAS:\n")
            f.write("-" * 30 + "\n")
            
            # Estatísticas do primeiro arquivo
            f.write("\nPRIMEIRO ARQUIVO (IMAGENS):\n")
            questoes_com_gabarito = sum(1 for q in questoes1 if q['has_answer'])
            f.write(f"  - Com gabarito: {questoes_com_gabarito}\n")
            f.write(f"  - Sem gabarito: {len(questoes1) - questoes_com_gabarito}\n")
            
            # Estatísticas do segundo arquivo
            f.write("\nSEGUNDO ARQUIVO (DOCX):\n")
            questoes_com_gabarito = sum(1 for q in questoes2 if q['has_answer'])
            f.write(f"  - Com gabarito: {questoes_com_gabarito}\n")
            f.write(f"  - Sem gabarito: {len(questoes2) - questoes_com_gabarito}\n")
            
            # Estatísticas do arquivo unificado
            f.write("\nARQUIVO UNIFICADO:\n")
            questoes_com_gabarito = sum(1 for q in questoes_unificadas if q['has_answer'])
            f.write(f"  - Com gabarito: {questoes_com_gabarito}\n")
            f.write(f"  - Sem gabarito: {len(questoes_unificadas) - questoes_com_gabarito}\n")
            
            f.write("\nDISTRIBUIÇÃO DAS FONTES:\n")
            fontes = {}
            for questao in questoes_unificadas:
                fonte = questao.get('fonte', 'Não especificada')
                fontes[fonte] = fontes.get(fonte, 0) + 1
            
            for fonte, count in fontes.items():
                f.write(f"  - {fonte}: {count} questões\n")
        
        print(f"📊 Relatório de unificação salvo em: {caminho_relatorio}")
    
    def unificar_arquivos(self, arquivo_json1, arquivo_json2, pasta_saida="questoes_unificadas"):
        """Unifica dois arquivos JSON em um único arquivo"""
        print("🔄 INICIANDO UNIFICAÇÃO DE ARQUIVOS JSON")
        print("=" * 50)
        
        # Criar pasta de saída
        os.makedirs(pasta_saida, exist_ok=True)
        
        # Carregar primeiro arquivo
        print(f"📁 Carregando {arquivo_json1}...")
        fixture1 = self.carregar_json(arquivo_json1)
        if not fixture1:
            return
        
        # Carregar segundo arquivo
        print(f"📁 Carregando {arquivo_json2}...")
        fixture2 = self.carregar_json(arquivo_json2)
        if not fixture2:
            return
        
        # Extrair questões de ambos os arquivos
        print("🔍 Extraindo questões do primeiro arquivo...")
        questoes1 = self.extrair_questoes_do_fixture(fixture1)
        print(f"   ✅ {len(questoes1)} questões extraídas")
        
        print("🔍 Extraindo questões do segundo arquivo...")
        questoes2 = self.extrair_questoes_do_fixture(fixture2)
        print(f"   ✅ {len(questoes2)} questões extraídas")
        
        # Unificar questões
        print("🔄 Unificando questões (removendo duplicatas)...")
        questoes_unificadas = self.unificar_questoes(questoes1, questoes2)
        print(f"   ✅ {len(questoes_unificadas)} questões após unificação")
        
        # Converter para formato Django
        print("📝 Convertendo para formato Django fixture...")
        fixture_unificado = self.converter_para_formato_django(questoes_unificadas)
        
        # Salvar arquivo unificado
        caminho_unificado = os.path.join(pasta_saida, "questions_fixture_unificado.json")
        with open(caminho_unificado, 'w', encoding='utf-8') as f:
            json.dump(fixture_unificado, f, ensure_ascii=False, indent=2)
        
        print(f"💾 Arquivo unificado salvo em: {caminho_unificado}")
        
        # Gerar relatório
        self.gerar_relatorio_unificacao(questoes1, questoes2, questoes_unificadas, pasta_saida)
        
        # Estatísticas finais
        print(f"\n✅ UNIFICAÇÃO CONCLUÍDA!")
        print(f"📊 Estatísticas:")
        print(f"   - Total de objetos: {len(fixture_unificado)}")
        print(f"   - Questions: {sum(1 for item in fixture_unificado if item['model'] == 'yourapp.Question')}")
        print(f"   - Alternatives: {sum(1 for item in fixture_unificado if item['model'] == 'yourapp.Alternative')}")
        print(f"   - CorrectAnswersSources: {sum(1 for item in fixture_unificado if item['model'] == 'yourapp.CorrectAnswersSources')}")
        print(f"📁 Pasta de saída: {pasta_saida}")

def main():
    """Função principal"""
    unificador = UnificadorQuestoes()
    
    # Configurar caminhos dos arquivos
    arquivo_json1 = "/home/yago/Tutorbots/questoes_processadas_copilot/questions_fixture.json"  # Arquivo das imagens
    arquivo_json2 = "/home/yago/Tutorbots/questoes_processadas_docx/questions_fixture.json"  # Arquivo do DOCX
    
    # Verificar se os arquivos existem
    if not os.path.exists(arquivo_json1):
        print(f"❌ Arquivo não encontrado: {arquivo_json1}")
        return
    
    if not os.path.exists(arquivo_json2):
        print(f"❌ Arquivo não encontrado: {arquivo_json2}")
        return
    
    # Executar unificação
    unificador.unificar_arquivos(arquivo_json1, arquivo_json2)

if __name__ == "__main__":
    main()
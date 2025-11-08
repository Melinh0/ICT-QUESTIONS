import pyautogui
import pyperclip
import time
import json
import re

# Configurações (altere conforme necessário)
COORDENADA_BARRA_ENDERECO = (264, 51)  # Coordenada da barra de endereço do navegador
COORDENADA_CTRL = (589, 367)   # Coordenada para clicar antes do Ctrl+A/C
COORDENADA_CLIQUE_LONGO = (672, 709)  # Coordenada para clique longo
COORDENADA_CLIQUE_NORMAL = (487, 632) # Coordenada para clique normal
REPETICOES_POR_LINK = 5
PASTA_SAIDA = r"C:\Users\Assistencia\Documents\GitHub\ICT-QUESTIONS\questões_CPRM"
ARQUIVO_SAIDA = f"{PASTA_SAIDA}\\questoes_organizadas.json"

# Lista de links para processar
LINKS = [
    "https://app.qconcursos.com/playground/questoes?discipline_ids%5B%5D=20&subject_ids%5B%5D=18060&subject_ids%5B%5D=18061&subject_ids%5B%5D=18065&subject_ids%5B%5D=18070&subject_ids%5B%5D=18071&subject_ids%5B%5D=18074&subject_ids%5B%5D=18076&subject_ids%5B%5D=18134&subject_ids%5B%5D=18135&subject_ids%5B%5D=18136&subject_ids%5B%5D=18137&subject_ids%5B%5D=18138&subject_ids%5B%5D=18139&course_sort_examining_board_id=63&sort=course_questions&per_page=25&limit=100",
    "https://app.qconcursos.com/playground/questoes?discipline_ids%5B%5D=2&subject_ids%5B%5D=15938&subject_ids%5B%5D=896&subject_ids%5B%5D=775&subject_ids%5B%5D=16349&subject_ids%5B%5D=1270&subject_ids%5B%5D=16167&subject_ids%5B%5D=1271&subject_ids%5B%5D=827&subject_ids%5B%5D=15943&subject_ids%5B%5D=15944&course_sort_examining_board_id=63&sort=course_questions&per_page=25&limit=100",
    "https://app.qconcursos.com/playground/questoes?discipline_ids%5B%5D=26&subject_ids%5B%5D=5858&subject_ids%5B%5D=5858&subject_ids%5B%5D=9668&subject_ids%5B%5D=4773&subject_ids%5B%5D=12440&subject_ids%5B%5D=12440&course_sort_examining_board_id=63&sort=course_questions&per_page=25&limit=100",
    "https://app.qconcursos.com/playground/questoes?discipline_ids%5B%5D=21&subject_ids%5B%5D=3446&subject_ids%5B%5D=19044&course_sort_examining_board_id=63&sort=course_questions&per_page=25&limit=100",
    "https://app.qconcursos.com/playground/questoes?discipline_ids%5B%5D=1&subject_ids%5B%5D=14621&subject_ids%5B%5D=14622&subject_ids%5B%5D=14624&subject_ids%5B%5D=15648&subject_ids%5B%5D=25201&subject_ids%5B%5D=25202&subject_ids%5B%5D=14625&subject_ids%5B%5D=16002&subject_ids%5B%5D=16392&subject_ids%5B%5D=14636&subject_ids%5B%5D=14628&subject_ids%5B%5D=14630&subject_ids%5B%5D=14631&subject_ids%5B%5D=14632&subject_ids%5B%5D=14633&subject_ids%5B%5D=16003&subject_ids%5B%5D=14635&subject_ids%5B%5D=16170&subject_ids%5B%5D=16175&subject_ids%5B%5D=16176&subject_ids%5B%5D=16177&subject_ids%5B%5D=16178&subject_ids%5B%5D=16179&subject_ids%5B%5D=16180&subject_ids%5B%5D=16181&subject_ids%5B%5D=16182&subject_ids%5B%5D=16183&subject_ids%5B%5D=16185&subject_ids%5B%5D=16174&subject_ids%5B%5D=16184&subject_ids%5B%5D=16157&subject_ids%5B%5D=16158&subject_ids%5B%5D=16171&subject_ids%5B%5D=20181&subject_ids%5B%5D=16159&subject_ids%5B%5D=16160&subject_ids%5B%5D=16161&subject_ids%5B%5D=16187&subject_ids%5B%5D=16172&subject_ids%5B%5D=16173&subject_ids%5B%5D=16006&subject_ids%5B%5D=16008&subject_ids%5B%5D=16009&subject_ids%5B%5D=20182&subject_ids%5B%5D=16010&subject_ids%5B%5D=14647&subject_ids%5B%5D=14648&subject_ids%5B%5D=14649&subject_ids%5B%5D=17753&subject_ids%5B%5D=14650&subject_ids%5B%5D=14651&subject_ids%5B%5D=14652&subject_ids%5B%5D=14653&subject_ids%5B%5D=14654&subject_ids%5B%5D=15724&subject_ids%5B%5D=15727&subject_ids%5B%5D=30912&subject_ids%5B%5D=14655&subject_ids%5B%5D=20185&subject_ids%5B%5D=14657&subject_ids%5B%5D=14659&subject_ids%5B%5D=14988&course_sort_examining_board_id=63&sort=course_questions&per_page=25&limit=100",
    "https://app.qconcursos.com/playground/questoes?discipline_ids%5B%5D=13&subject_ids%5B%5D=14553&subject_ids%5B%5D=20313&subject_ids%5B%5D=20317&subject_ids%5B%5D=20318&subject_ids%5B%5D=20321&subject_ids%5B%5D=20322&subject_ids%5B%5D=21769&subject_ids%5B%5D=21769&course_sort_examining_board_id=63&sort=course_questions&per_page=25&limit=100",
    "https://app.qconcursos.com/playground/questoes?discipline_ids%5B%5D=4&subject_ids%5B%5D=15441&subject_ids%5B%5D=15443&subject_ids%5B%5D=15442&subject_ids%5B%5D=15446&subject_ids%5B%5D=17777&subject_ids%5B%5D=17778&subject_ids%5B%5D=26412&subject_ids%5B%5D=26412&subject_ids%5B%5D=15453&subject_ids%5B%5D=15453&subject_ids%5B%5D=15463&subject_ids%5B%5D=15463&subject_ids%5B%5D=15459&subject_ids%5B%5D=15459&subject_ids%5B%5D=15460&subject_ids%5B%5D=15460&subject_ids%5B%5D=15461&subject_ids%5B%5D=15461&subject_ids%5B%5D=16021&subject_ids%5B%5D=16021&course_sort_examining_board_id=63&sort=course_questions&per_page=25&limit=100",
    "https://app.qconcursos.com/playground/questoes?discipline_ids%5B%5D=56&subject_ids%5B%5D=32473&subject_ids%5B%5D=32473&subject_ids%5B%5D=31678&subject_ids%5B%5D=31678&subject_ids%5B%5D=31000&subject_ids%5B%5D=31000&subject_ids%5B%5D=28062&subject_ids%5B%5D=28062&subject_ids%5B%5D=16798&subject_ids%5B%5D=16799&subject_ids%5B%5D=8323&subject_ids%5B%5D=16801&subject_ids%5B%5D=4572&subject_ids%5B%5D=16802&subject_ids%5B%5D=16806&subject_ids%5B%5D=16806&subject_ids%5B%5D=16803&course_sort_examining_board_id=63&sort=course_questions&per_page=25&limit=100"
]

# Lista para armazenar todas as questões
todas_questoes = []

def tem_imagem(texto):
    """Verifica se o texto contém referências a imagens"""
    padroes_imagem = [
        r'\.(png|jpg|jpeg|gif|bmp|svg)',
        r'Captura_de tela',
        r'Captura de tela',
        r'figura abaixo',
        r'imagem abaixo',
        r'gráfico abaixo',
        r'\(.*\d+×\d+.*\)'
    ]
    
    for padrao in padroes_imagem:
        if re.search(padrao, texto, re.IGNORECASE):
            return True
    return False

def limpar_texto(texto):
    """Remove texto indesejado como gabarito comentado e estatísticas"""
    texto = re.sub(r'Gabarito comentado.*', '', texto, flags=re.DOTALL)
    texto = re.sub(r'Comentários de alunos.*', '', texto, flags=re.DOTALL)
    texto = re.sub(r'\d+\s*Estatísticas.*', '', texto)
    texto = re.sub(r'Aulas\s*\d+.*', '', texto)
    texto = re.sub(r'\s+', ' ', texto)
    return texto.strip()

def extrair_cargo_completo(enunciado_completo):
    """Extrai o cargo completo do enunciado usando uma abordagem mais robusta"""
    padroes_cargo = [
        r'^([A-Z][a-záàâãéèêíïóôõöúçñ\s\-]+?)(?=\s+[A-Z][a-z]{2,}\s+)',
        r'^([A-Z][a-záàâãéèêíïóôõöúçñ\s\-]+?)(?=\s+(?:Em|Em\s+uma|Em\s+um|No|Na|Os|As|Um|Uma))',
        r'^((?:[A-Z][a-záàâãéèêíïóôõöúçñ]+\s*){1,5})(?=\s*[A-Z][a-z]{3,})',
        r'^([A-Z][a-záàâãéèêíïóôõöúçñ\s\-]+?)(?=\s*[•\d])',
    ]
    
    palavras_cargo = [
        'administrativo', 'analista', 'técnico', 'assistente', 'coordenador', 
        'diretor', 'gerente', 'agente', 'fiscal', 'especialista', 'professor',
        'auditor', 'contador', 'engenheiro', 'advogado', 'médico', 'enfermeiro',
        'psicólogo', 'assistente social', 'secretário', 'administrador'
    ]
    
    for padrao in padroes_cargo:
        match = re.search(padrao, enunciado_completo)
        if match:
            cargo_candidato = match.group(1).strip()
            
            cargo_lower = cargo_candidato.lower()
            if any(palavra in cargo_lower for palavra in palavras_cargo):
                
                padroes_extensao = [
                    r'^([A-Z][a-záàâãéèêíïóôõöúçñ\s\-]+?(?:\s+[A-Z][a-záàâãéèêíïóôõöúçñ\s\-]+){0,3})',
                    r'^(.+?)(?=\s+(?:Em\s+|No\s+|Na\s+|Os\s+|As\s+|Um\s+|Uma\s+))',
                ]
                
                for padrao_ext in padroes_extensao:
                    match_ext = re.search(padrao_ext, enunciado_completo)
                    if match_ext:
                        cargo_extendido = match_ext.group(1).strip()
                        cargo_ext_lower = cargo_extendido.lower()
                        
                        palavras_encontradas = sum(1 for palavra in palavras_cargo if palavra in cargo_ext_lower)
                        if palavras_encontradas >= 1 and len(cargo_extendido) > len(cargo_candidato):
                            cargo_candidato = cargo_extendido
                
                enunciado_limpo = enunciado_completo[len(cargo_candidato):].strip()
                return cargo_candidato, enunciado_limpo
    
    palavras = enunciado_completo.split()
    if len(palavras) >= 2:
        for i in range(1, min(5, len(palavras))):
            candidato = ' '.join(palavras[:i])
            candidato_lower = candidato.lower()
            
            if any(palavra in candidato_lower for palavra in palavras_cargo):
                enunciado_restante = ' '.join(palavras[i:])
                return candidato, enunciado_restante
    
    return None, enunciado_completo

def parse_questoes(texto):
    """Extrai questões do texto copiado e retorna lista organizada"""
    questoes = []
    
    if tem_imagem(texto):
        print("  -> Texto contém imagens, ignorando...")
        return []
    
    padrao_questao = re.compile(r'^(Q\d+)\s*$', re.MULTILINE)
    blocos = padrao_questao.split(texto)[1:]
    
    for i in range(0, len(blocos), 2):
        if i + 1 >= len(blocos):
            break
            
        codigo = blocos[i].strip()
        conteudo = blocos[i + 1]
        
        if tem_imagem(conteudo):
            print(f"  -> Questão {codigo} contém imagens, ignorando...")
            continue
        
        questao = extrair_informacoes_questao(codigo, conteudo)
        if questao:
            questoes.append(questao)
    
    return questoes

def extrair_informacoes_questao(codigo, texto):
    """Extrai informações específicas de cada questão"""
    try:
        linhas = texto.strip().split('\n')
        linhas = [linha.strip() for linha in linhas if linha.strip()]
        
        banca_idx = next((i for i, linha in enumerate(linhas) if 'FGV' in linha), -1)
        tema_idx = banca_idx + 1 if banca_idx != -1 else -1
        assunto_idx = tema_idx + 1 if tema_idx != -1 else -1
        concurso_idx = assunto_idx + 1 if assunto_idx != -1 else -1
        
        if banca_idx == -1:
            return None
            
        banca = linhas[banca_idx]
        tema = linhas[tema_idx] if tema_idx < len(linhas) else "Não informado"
        assunto = linhas[assunto_idx] if assunto_idx < len(linhas) else "Não informado"
        concurso = linhas[concurso_idx] if concurso_idx < len(linhas) else "Não informado"
        
        inicio_enunciado = concurso_idx + 1
        if inicio_enunciado >= len(linhas):
            return None
            
        inicio_alternativas = -1
        for i in range(inicio_enunciado, len(linhas)):
            if re.match(r'^A\s*$', linhas[i]) or re.match(r'^A\.', linhas[i]):
                inicio_alternativas = i
                break
        
        if inicio_alternativas == -1:
            return None
            
        enunciado_completo = ' '.join(linhas[inicio_enunciado:inicio_alternativas])
        
        cargo, enunciado_limpo = extrair_cargo_completo(enunciado_completo)
        
        alternativas = {}
        i = inicio_alternativas
        
        while i < len(linhas):
            linha = linhas[i]
            
            if re.match(r'^[A-E]\s*$', linha) or re.match(r'^[A-E]\.', linha):
                letra_alternativa = linha[0]
                
                texto_alternativa = []
                j = i + 1
                while j < len(linhas) and not re.match(r'^[A-E]\s*$', linhas[j]) and not re.match(r'^[A-E]\.', linhas[j]):
                    texto_alternativa.append(linhas[j])
                    j += 1
                
                texto_completo = ' '.join(texto_alternativa)
                texto_completo = limpar_texto(texto_completo)
                
                alternativas[letra_alternativa] = texto_completo
                i = j - 1
            
            i += 1
        
        if not alternativas:
            return None
        
        questao = {
            "codigo": codigo,
            "banca": banca,
            "tema": tema,
            "assunto": assunto,
            "concurso": concurso,
            "enunciado": limpar_texto(enunciado_limpo)
        }
        
        if cargo:
            questao["cargo"] = cargo
        
        questao["alternativas"] = alternativas
        
        return questao
    
    except Exception as e:
        print(f"Erro ao processar questão {codigo}: {e}")
        return None

def navegar_para_link(link):
    """Navega para o link especificado no navegador"""
    print(f"\nNavegando para: {link}")
    
    # Clica na barra de endereço
    pyautogui.click(COORDENADA_BARRA_ENDERECO)
    time.sleep(1)
    
    # Seleciona todo o texto atual (Ctrl+A)
    pyautogui.hotkey('ctrl', 'a')
    time.sleep(0.5)
    
    # Cola o novo link
    pyperclip.copy(link)
    pyautogui.hotkey('ctrl', 'v')
    time.sleep(0.5)
    
    # Pressiona Enter para navegar
    pyautogui.press('enter')
    time.sleep(7)  # Aguarda o carregamento da página

def executar_automacao_por_link():
    """Executa a automação para o link atual"""
    try:
        for i in range(REPETICOES_POR_LINK):
            print(f"Executando ciclo {i+1} de {REPETICOES_POR_LINK}...")
            
            # Clica na coordenada do campo de texto
            pyautogui.click(COORDENADA_CTRL)
            time.sleep(1)
            
            # Simula Ctrl+A e Ctrl+C
            pyautogui.hotkey('ctrl', 'a')
            time.sleep(0.5)
            pyautogui.hotkey('ctrl', 'c')
            time.sleep(1)
            
            # Obtém conteúdo copiado
            texto_copiado = pyperclip.paste()
            
            if texto_copiado:
                print(f"Processando {len(texto_copiado)} caracteres copiados...")
                
                if tem_imagem(texto_copiado):
                    print("  -> Conteúdo copiado contém imagens, ignorando ciclo...")
                else:
                    questões = parse_questoes(texto_copiado)
                    todas_questoes.extend(questões)
                    print(f"Encontradas {len(questões)} questões no ciclo {i+1}")
            else:
                print("Nenhum conteúdo copiado encontrado")
            
            # Clique longo (5 segundos)
            pyautogui.mouseDown(COORDENADA_CLIQUE_LONGO)
            time.sleep(5)
            pyautogui.mouseUp()
            
            # Clique normal
            pyautogui.click(COORDENADA_CLIQUE_NORMAL)
            
            # Intervalo entre ciclos
            time.sleep(2)
            
    except Exception as e:
        print(f"Erro durante automação: {e}")

def processar_todos_links():
    """Processa todos os links da lista"""
    print(f"Iniciando processamento de {len(LINKS)} links...")
    
    for i, link in enumerate(LINKS, 1):
        print(f"\n{'='*50}")
        print(f"Processando link {i} de {len(LINKS)}")
        print(f"{'='*50}")
        
        # Navega para o link
        navegar_para_link(link)
        
        # Executa a automação para o link atual
        executar_automacao_por_link()
        
        print(f"Concluído link {i} de {len(LINKS)}")
        
        # Pequena pausa entre links
        time.sleep(2)

def salvar_questoes_json():
    """Salva as questões no arquivo JSON"""
    try:
        # Remove questões None
        questoes_validas = [q for q in todas_questoes if q is not None]
        
        with open(ARQUIVO_SAIDA, 'w', encoding='utf-8') as f:
            json.dump(questoes_validas, f, ensure_ascii=False, indent=2)
        print(f"\nTotal de {len(questoes_validas)} questões válidas salvas em {ARQUIVO_SAIDA}")
        
        # Mostra estatísticas detalhadas
        if questoes_validas:
            cargos_com_cargo = [q for q in questoes_validas if 'cargo' in q]
            print(f"Questões com cargo identificado: {len(cargos_com_cargo)}")
            
            if cargos_com_cargo:
                print("Exemplos de cargos extraídos:")
                cargos_unicos = list(set(q['cargo'] for q in cargos_com_cargo))
                for cargo in cargos_unicos[:5]:
                    print(f"  - {cargo}")
            
    except Exception as e:
        print(f"Erro ao salvar arquivo JSON: {e}")

if __name__ == "__main__":
    print("Iniciando bot de automação para múltiplos links...")
    print("Certifique-se de que o navegador está aberto e visível")
    print("Posicione o cursor e aguarde 5 segundos...")
    time.sleep(5)
    
    processar_todos_links()
    salvar_questoes_json()
    
    print("\nProcesso concluído! Todos os links foram processados.")
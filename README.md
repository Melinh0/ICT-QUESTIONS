# ICT-QUESTIONS

Sistema completo para processamento e gerenciamento de questões de múltipla escolha para estudos de certificação HCIA Computing.

## 📋 Descrição do Projeto

Este projeto consiste em uma aplicação desktop para estudo de questões de certificação, com funcionalidades de:

- **Processamento automático** de documentos DOCX com questões
- **Reconhecimento de texto** usando Microsoft Copilot
- **Interface gráfica** para estudo e revisão
- **Sistema de unificação** de bancos de questões
- **Exportação** para formato Django fixture

## 🛠️ Configuração do Ambiente

### Pré-requisitos

- Python 3.8 ou superior
- Git
- Navegador Microsoft Edge (para integração com Copilot)

### 1. Clonar o Repositório

```bash
git clone https://github.com/seu-usuario/ICT-QUESTIONS.git
cd ICT-QUESTIONS
```

### 2. Criar Ambiente Virtual

**Windows:**
```bash
python -m venv venv
venv\Scripts\activate
```

**Linux/MacOS:**
```bash
python3 -m venv venv
source venv/bin/activate
```

### 3. Instalar Dependências

```bash
pip install -r requirements.txt
```

### 4. Estrutura do Projeto

```
ICT-QUESTIONS/
├── venv/                          # Ambiente virtual
├── questoes_processadas_copilot/  # Questões processadas por imagens
├── questoes_processadas_docx/     # Questões processadas por texto
├── questoes_unificadas/           # Banco de questões unificado
├── study_app.py                   # Aplicação principal de estudo
├── extractor_copilot_images.py    # Extrator via imagens
├── extractor_docx_text.py         # Extrator via texto DOCX
├── unificador_questoes.py         # Unificador de bancos
├── requirements.txt               # Dependências do projeto
└── README.md                      # Este arquivo
```

## 📦 Dependências do Projeto

O arquivo `requirements.txt` contém:

```txt
tkinter
pillow
deep-translator
pyautogui
pyperclip
python-docx
json
datetime
glob
re
os
sys
time
```

## 🚀 Como Usar

### 1. Processamento de Documentos

**Processar via Imagens (Copilot):**
```bash
python extractor_copilot_images.py
```

**Processar via Texto DOCX:**
```bash
python extractor_docx_text.py
```

### 2. Unificar Bancos de Questões

```bash
python unificador_questoes.py
```

### 3. Executar Aplicação de Estudo

```bash
python study_app.py
```

## ⚙️ Configuração do Copilot

Para usar o processamento automático com Copilot:

1. Tenha o Microsoft Edge instalado
2. Faça login com uma conta Microsoft
3. Acesse o Copilot (copilot.microsoft.com)
4. Posicione a janela do navegador conforme as coordenadas configuradas

## 🎯 Funcionalidades da Aplicação

### Interface de Estudo
- Navegação entre questões
- Alternância de idioma (Português/Inglês)
- Exibição de respostas e explicações
- Informações sobre nível e track das questões

### Processamento Automático
- Extração de texto de documentos DOCX
- Reconhecimento de imagens com Copilot
- Organização automática de questões
- Obtenção de gabaritos via IA

### Gerenciamento de Dados
- Unificação de múltiplos bancos
- Remoção de duplicatas
- Exportação para JSON/Django
- Relatórios detalhados

## 🔧 Solução de Problemas

### Erros Comuns

**ImportError: No module named 'tkinter'**
```bash
# Ubuntu/Debian
sudo apt-get install python3-tk

# Windows (usando Python oficial)
# O Tkinter já vem incluído
```

**PyAutoGUI - Permissão de Tela (MacOS)**
- Acesse: System Preferences → Security & Privacy → Privacy → Accessibility
- Adicione o Terminal ou seu IDE à lista

**Copilot não responde**
- Verifique se está logado no Microsoft Edge
- Confirme se o Copilot está acessível
- Ajuste as coordenadas em `extractor_copilot_images.py`

### Ajuste de Coordenadas

Se a automação não funcionar, edite as coordenadas no código:

```python
self.coordenadas = {
    'campo_prompt': (x, y),  # Ajuste conforme sua tela
    'enviar_mensagem': (x, y),
    # ... outras coordenadas
}
```

## 📊 Estrutura de Dados

### Formato das Questões
```json
{
    "enunciado": "Texto da pergunta",
    "alternativas": {
        "A": "Texto alternativa A",
        "B": "Texto alternativa B",
        "C": "Texto alternativa C", 
        "D": "Texto alternativa D"
    },
    "item_correto": "A",
    "explicacao": "Explicação detalhada",
    "fonte": "Fonte da resposta"
}
```

### Formato Django Fixture
```json
[
    {
        "model": "yourapp.Question",
        "pk": 1,
        "fields": {
            "text": "Enunciado...",
            "level": "HCIA",
            "has_answer": true
        }
    }
]
```

## 🤝 Contribuição

1. Fork o projeto
2. Crie uma branch para sua feature (`git checkout -b feature/AmazingFeature`)
3. Commit suas mudanças (`git commit -m 'Add some AmazingFeature'`)
4. Push para a branch (`git push origin feature/AmazingFeature`)
5. Abra um Pull Request

## 🆕 Atualizações Futuras

- [ ] Integração com API oficial do Copilot
- [ ] Sistema de revisão espaçada
- [ ] Exportação para Anki
- [ ] Estatísticas de desempenho
- [ ] Modo simulado de prova

## 📞 Suporte

Em caso de problemas:

1. Verifique se todas as dependências foram instaladas
2. Confirme que o ambiente virtual está ativado
3. Execute `python --version` para verificar a versão do Python
4. Consulte a seção de Solução de Problemas acima

---

**Desenvolvido para estudos de certificação HCIA Computing** 🎓
#!/usr/bin/env python3
# -*- coding: utf-8 -*-

import json
import tkinter as tk
from tkinter import ttk, messagebox
import os
from deep_translator import GoogleTranslator

class StudyApp:
    def __init__(self, root):
        self.root = root
        self.root.title("Study Questions - HCIA Computing")
        self.root.geometry("1000x700")
        self.root.configure(bg='#f0f0f0')
        
        # Variáveis
        self.questions = []
        self.current_question_index = 0
        self.current_language = "english"  # ou "portuguese"
        self.selected_alternative = tk.StringVar()  # CORREÇÃO: Adicionar esta linha
        self.translator = GoogleTranslator(source='auto', target='en')
        
        # Carregar questões
        self.load_questions()
        
        # Criar interface
        self.create_widgets()
        self.show_question()
    
    def load_questions(self):
        """Carrega as questões do arquivo JSON"""
        try:
            json_path = "/home/yago/ICT-QUESTIONS/questoes_unificadas/questions_fixture_unificado.json"
            
            if not os.path.exists(json_path):
                messagebox.showerror("Erro", f"Arquivo não encontrado: {json_path}")
                return
            
            with open(json_path, 'r', encoding='utf-8') as f:
                data = json.load(f)
            
            # Extrair questões do formato Django fixture
            self.questions = self.extract_questions_from_fixture(data)
            
            if not self.questions:
                messagebox.showerror("Erro", "Nenhuma questão encontrada no arquivo")
                return
                
            messagebox.showinfo("Sucesso", f"Carregadas {len(self.questions)} questões")
            
        except Exception as e:
            messagebox.showerror("Erro", f"Erro ao carregar questões: {str(e)}")
    
    def extract_questions_from_fixture(self, fixture_data):
        """Extrai questões do formato Django fixture"""
        questions_dict = {}
        
        # Primeiro, coletar todas as questões
        for item in fixture_data:
            if item['model'] == 'yourapp.Question':
                question_id = item['pk']
                questions_dict[question_id] = {
                    'id': question_id,
                    'enunciado': item['fields']['text'],
                    'alternativas': {},
                    'item_correto': None,
                    'explicacao': '',
                    'fonte': '',
                    'has_answer': item['fields']['has_answer'],
                    'level': item['fields']['level'],
                    'track': item['fields']['track']
                }
        
        # Coletar alternativas
        for item in fixture_data:
            if item['model'] == 'yourapp.Alternative':
                question_id = item['fields']['question']
                alternativa_id = item['pk']
                letra = self.get_alternative_letter(alternativa_id, fixture_data)
                
                if question_id in questions_dict:
                    questions_dict[question_id]['alternativas'][letra] = {
                        'texto': item['fields']['text'],
                        'is_correct': item['fields']['is_correct']
                    }
                    
                    # Marcar como correta
                    if item['fields']['is_correct']:
                        questions_dict[question_id]['item_correto'] = letra
        
        # Coletar fontes e explicações
        for item in fixture_data:
            if item['model'] == 'yourapp.CorrectAnswersSources':
                alternativa_id = item['fields']['alternative']
                fonte = item['fields']['source']
                
                # Encontrar a questão correspondente
                for questao in questions_dict.values():
                    if self.is_alternative_in_question(alternativa_id, questao['id'], fixture_data):
                        questao['fonte'] = fonte
                        # Se a fonte indica resposta do Copilot, usar como explicação
                        if "Copilot" in fonte:
                            questao['explicacao'] = fonte
                        break
        
        return list(questions_dict.values())
    
    def get_alternative_letter(self, alternativa_id, fixture_data):
        """Obtém a letra da alternativa baseado na ordem"""
        # Encontrar questão
        questao_id = None
        for item in fixture_data:
            if item['model'] == 'yourapp.Alternative' and item['pk'] == alternativa_id:
                questao_id = item['fields']['question']
                break
        
        if not questao_id:
            return 'A'
        
        # Coletar todas as alternativas da questão
        alternativas = []
        for item in fixture_data:
            if item['model'] == 'yourapp.Alternative' and item['fields']['question'] == questao_id:
                alternativas.append((item['pk'], item['fields']['text']))
        
        alternativas.sort(key=lambda x: x[0])
        
        # Mapear para letras
        letras = ['A', 'B', 'C', 'D', 'E']
        for i, (alt_id, texto) in enumerate(alternativas):
            if alt_id == alternativa_id:
                return letras[i] if i < len(letras) else 'A'
        
        return 'A'
    
    def is_alternative_in_question(self, alternativa_id, questao_id, fixture_data):
        """Verifica se uma alternativa pertence a uma questão"""
        for item in fixture_data:
            if (item['model'] == 'yourapp.Alternative' and 
                item['pk'] == alternativa_id and 
                item['fields']['question'] == questao_id):
                return True
        return False
    
    def create_widgets(self):
        """Cria os elementos da interface"""
        # Frame principal
        main_frame = ttk.Frame(self.root, padding="10")
        main_frame.grid(row=0, column=0, sticky=(tk.W, tk.E, tk.N, tk.S))
        
        # Configurar grid
        self.root.columnconfigure(0, weight=1)
        self.root.rowconfigure(0, weight=1)
        main_frame.columnconfigure(1, weight=1)
        
        # Controles superiores
        controls_frame = ttk.Frame(main_frame)
        controls_frame.grid(row=0, column=0, columnspan=2, sticky=(tk.W, tk.E), pady=(0, 10))
        
        # Botão de idioma
        self.language_btn = ttk.Button(
            controls_frame, 
            text="Português", 
            command=self.toggle_language
        )
        self.language_btn.grid(row=0, column=0, padx=(0, 10))
        
        # Navegação
        nav_frame = ttk.Frame(controls_frame)
        nav_frame.grid(row=0, column=1)
        
        ttk.Button(nav_frame, text="← Anterior", command=self.previous_question).grid(row=0, column=0, padx=(0, 5))
        ttk.Button(nav_frame, text="Próximo →", command=self.next_question).grid(row=0, column=1, padx=(5, 0))
        
        # Informações da questão
        info_frame = ttk.Frame(controls_frame)
        info_frame.grid(row=0, column=2, padx=(20, 0))
        
        self.question_info = ttk.Label(
            info_frame, 
            text="Questão 1/150 | HCIA | Computing",
            font=('Arial', 10)
        )
        self.question_info.grid(row=0, column=0)
        
        # Área da questão
        question_frame = ttk.LabelFrame(main_frame, text="QUESTÃO", padding="10")
        question_frame.grid(row=1, column=0, columnspan=2, sticky=(tk.W, tk.E, tk.N, tk.S), pady=(0, 10))
        question_frame.columnconfigure(0, weight=1)
        
        # Enunciado
        self.question_text = tk.Text(
            question_frame, 
            height=6, 
            wrap=tk.WORD, 
            font=('Arial', 11),
            bg='white'
        )
        self.question_text.grid(row=0, column=0, sticky=(tk.W, tk.E), pady=(0, 10))
        
        # Scrollbar para o enunciado
        scrollbar = ttk.Scrollbar(question_frame, orient=tk.VERTICAL, command=self.question_text.yview)
        scrollbar.grid(row=0, column=1, sticky=(tk.N, tk.S))
        self.question_text.configure(yscrollcommand=scrollbar.set)
        
        # Alternativas
        alternatives_frame = ttk.LabelFrame(main_frame, text="ALTERNATIVAS", padding="10")
        alternatives_frame.grid(row=2, column=0, columnspan=2, sticky=(tk.W, tk.E), pady=(0, 10))
        alternatives_frame.columnconfigure(0, weight=1)
        
        self.alternative_widgets = []
        
        for i in range(5):
            frame = ttk.Frame(alternatives_frame)
            frame.grid(row=i, column=0, sticky=(tk.W, tk.E), pady=2)
            frame.columnconfigure(1, weight=1)
            
            # Label com a letra da alternativa
            label = ttk.Label(frame, text=f"{chr(65+i)})", font=('Arial', 10, 'bold'))
            label.grid(row=0, column=0, padx=(0, 10), sticky=tk.W)
            
            text_widget = tk.Text(
                frame, 
                height=2, 
                wrap=tk.WORD, 
                font=('Arial', 10),
                bg='white',
                width=60
            )
            text_widget.grid(row=0, column=1, sticky=(tk.W, tk.E))
            
            self.alternative_widgets.append(text_widget)
        
        # Botão de ver resposta
        ttk.Button(
            main_frame, 
            text="Mostrar Resposta", 
            command=self.show_answer
        ).grid(row=3, column=0, columnspan=2, pady=(10, 5))
        
        # Área de resposta
        answer_frame = ttk.LabelFrame(main_frame, text="RESPOSTA E EXPLICAÇÃO", padding="10")
        answer_frame.grid(row=4, column=0, columnspan=2, sticky=(tk.W, tk.E, tk.N, tk.S))
        answer_frame.columnconfigure(0, weight=1)
        answer_frame.rowconfigure(0, weight=1)
        
        self.answer_text = tk.Text(
            answer_frame, 
            height=8, 
            wrap=tk.WORD, 
            font=('Arial', 10),
            bg='#f8f8f8'
        )
        self.answer_text.grid(row=0, column=0, sticky=(tk.W, tk.E, tk.N, tk.S))
        
        # Scrollbar para resposta
        answer_scrollbar = ttk.Scrollbar(answer_frame, orient=tk.VERTICAL, command=self.answer_text.yview)
        answer_scrollbar.grid(row=0, column=1, sticky=(tk.N, tk.S))
        self.answer_text.configure(yscrollcommand=answer_scrollbar.set)
        
        # Configurar expansão
        main_frame.rowconfigure(4, weight=1)
    
    def show_question(self):
        """Exibe a questão atual"""
        if not self.questions:
            return
        
        question = self.questions[self.current_question_index]
        
        # Atualizar informações
        self.question_info.config(
            text=f"Questão {self.current_question_index + 1}/{len(self.questions)} | {question['level']} | {question['track']}"
        )
        
        # Exibir enunciado
        self.question_text.config(state=tk.NORMAL)
        self.question_text.delete(1.0, tk.END)
        self.question_text.insert(1.0, question['enunciado'])
        self.question_text.config(state=tk.DISABLED)
        
        # Exibir alternativas
        for i, letra in enumerate(['A', 'B', 'C', 'D', 'E']):
            text_widget = self.alternative_widgets[i]
            text_widget.config(state=tk.NORMAL)
            text_widget.delete(1.0, tk.END)
            
            if letra in question['alternativas']:
                text_widget.insert(1.0, question['alternativas'][letra]['texto'])
                text_widget.config(state=tk.DISABLED)
            else:
                text_widget.insert(1.0, "")
                text_widget.config(state=tk.DISABLED)
        
        # Limpar área de resposta
        self.answer_text.config(state=tk.NORMAL)
        self.answer_text.delete(1.0, tk.END)
        self.answer_text.config(state=tk.DISABLED)
    
    def show_answer(self):
        """Mostra a resposta correta e explicação"""
        if not self.questions:
            return
        
        question = self.questions[self.current_question_index]
        
        if not question['has_answer']:
            self.answer_text.config(state=tk.NORMAL)
            self.answer_text.delete(1.0, tk.END)
            self.answer_text.insert(1.0, "Esta questão não possui gabarito definido.")
            self.answer_text.config(state=tk.DISABLED)
            return
        
        resposta_texto = f"RESPOSTA CORRETA: {question['item_correto']}\n\n"
        
        if question['explicacao']:
            resposta_texto += f"EXPLICAÇÃO:\n{question['explicacao']}\n\n"
        
        if question['fonte']:
            resposta_texto += f"FONTE: {question['fonte']}"
        
        self.answer_text.config(state=tk.NORMAL)
        self.answer_text.delete(1.0, tk.END)
        self.answer_text.insert(1.0, resposta_texto)
        self.answer_text.config(state=tk.DISABLED)
    
    def next_question(self):
        """Vai para a próxima questão"""
        if self.current_question_index < len(self.questions) - 1:
            self.current_question_index += 1
            self.show_question()
    
    def previous_question(self):
        """Volta para a questão anterior"""
        if self.current_question_index > 0:
            self.current_question_index -= 1
            self.show_question()
    
    def toggle_language(self):
        """Alterna entre inglês e português"""
        if self.current_language == "english":
            self.current_language = "portuguese"
            self.language_btn.config(text="English")
            self.translate_to_portuguese()
        else:
            self.current_language = "english"
            self.language_btn.config(text="Português")
            self.show_question()  # Volta ao original
    
    def translate_to_portuguese(self):
        """Traduz a questão atual para português"""
        if not self.questions:
            return
        
        question = self.questions[self.current_question_index]
        
        try:
            # Configurar tradutor para português
            self.translator.target = 'pt'
            
            # Traduzir enunciado
            translated_question = self.translator.translate(question['enunciado'])
            self.question_text.config(state=tk.NORMAL)
            self.question_text.delete(1.0, tk.END)
            self.question_text.insert(1.0, translated_question)
            self.question_text.config(state=tk.DISABLED)
            
            # Traduzir alternativas
            for i, letra in enumerate(['A', 'B', 'C', 'D', 'E']):
                if letra in question['alternativas']:
                    translated_alt = self.translator.translate(question['alternativas'][letra]['texto'])
                    text_widget = self.alternative_widgets[i]
                    text_widget.config(state=tk.NORMAL)
                    text_widget.delete(1.0, tk.END)
                    text_widget.insert(1.0, translated_alt)
                    text_widget.config(state=tk.DISABLED)
            
            # Traduzir explicação se existir
            if question['explicacao']:
                translated_explanation = self.translator.translate(question['explicacao'])
                self.answer_text.config(state=tk.NORMAL)
                self.answer_text.delete(1.0, tk.END)
                
                resposta_texto = f"RESPOSTA CORRETA: {question['item_correto']}\n\n"
                resposta_texto += f"EXPLICAÇÃO:\n{translated_explanation}\n\n"
                
                if question['fonte']:
                    resposta_texto += f"FONTE: {question['fonte']}"
                
                self.answer_text.insert(1.0, resposta_texto)
                self.answer_text.config(state=tk.DISABLED)
                
        except Exception as e:
            messagebox.showerror("Erro de Tradução", f"Erro ao traduzir: {str(e)}")
        finally:
            # Restaurar tradutor para inglês
            self.translator.target = 'en'

def main():
    root = tk.Tk()
    app = StudyApp(root)
    root.mainloop()

if __name__ == "__main__":
    main()
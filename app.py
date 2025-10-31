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
        self.current_language = "english"  # Começa em inglês por padrão
        self.translated_questions = {}  # Cache para questões traduzidas
        self.translator = GoogleTranslator(source='auto', target='en')
        
        # Carregar questões
        self.load_questions()
        
        # Criar interface
        self.create_widgets()
        self.show_question()
    
    def load_questions(self):
        """Carrega as questões do arquivo JSON"""
        try:
            json_path = "/home/yago/ICT-QUESTIONS/questoes_explicadas/explicacoes_completas.json"
            
            if not os.path.exists(json_path):
                messagebox.showerror("Erro", f"Arquivo não encontrado: {json_path}")
                return
            
            with open(json_path, 'r', encoding='utf-8') as f:
                data = json.load(f)
            
            # Extrair questões do formato
            self.questions = self.extract_questions_from_json(data)
            
            if not self.questions:
                messagebox.showerror("Erro", "Nenhuma questão encontrada no arquivo")
                return
                
            messagebox.showinfo("Sucesso", f"Carregadas {len(self.questions)} questões")
            
        except Exception as e:
            messagebox.showerror("Erro", f"Erro ao carregar questões: {str(e)}")
    
    def extract_questions_from_json(self, json_data):
        """Extrai questões do formato JSON fornecido"""
        questions_dict = {}
        
        # Primeiro, coletar todas as questões
        for item in json_data:
            if item['model'] == 'questions.Question':
                question_id = item['pk']
                questions_dict[question_id] = {
                    'id': question_id,
                    'text_en': item['fields']['text'],  # Texto original em inglês
                    'explanation_pt': item['fields']['explicacao'],  # Explicação original em português
                    'alternatives': {},
                    'correct_answer': None,
                    'level': item['fields']['level'],
                    'track': item['fields']['track'],
                    'has_answer': item['fields']['has_answer'],
                    'weight': item['fields']['weight']
                }
        
        # Coletar alternativas
        for item in json_data:
            if item['model'] == 'alternatives.Alternative':
                question_id = item['fields']['question']
                alternative_id = item['pk']
                
                if question_id in questions_dict:
                    # Determinar a letra da alternativa (A, B, C, D)
                    alternatives_count = len(questions_dict[question_id]['alternatives'])
                    letter = chr(65 + alternatives_count)  # A, B, C, D
                    
                    questions_dict[question_id]['alternatives'][letter] = {
                        'text': item['fields']['text'],
                        'is_correct': item['fields']['is_correct']
                    }
                    
                    # Marcar como correta
                    if item['fields']['is_correct']:
                        questions_dict[question_id]['correct_answer'] = letter
        
        # Converter para lista ordenada por ID
        questions_list = [questions_dict[key] for key in sorted(questions_dict.keys())]
        return questions_list
    
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
        
        # Botão de idioma - Inicia em "Português" pois o idioma atual é inglês
        self.language_btn = ttk.Button(
            controls_frame, 
            text="Português", 
            command=self.toggle_language
        )
        self.language_btn.grid(row=0, column=0, padx=(0, 10))
        
        # Navegação
        nav_frame = ttk.Frame(controls_frame)
        nav_frame.grid(row=0, column=1)
        
        self.prev_btn = ttk.Button(nav_frame, text="← Previous", command=self.previous_question)
        self.prev_btn.grid(row=0, column=0, padx=(0, 5))
        
        self.next_btn = ttk.Button(nav_frame, text="Next →", command=self.next_question)
        self.next_btn.grid(row=0, column=1, padx=(5, 0))
        
        # Informações da questão
        info_frame = ttk.Frame(controls_frame)
        info_frame.grid(row=0, column=2, padx=(20, 0))
        
        self.question_info = ttk.Label(
            info_frame, 
            text="Question 1/150 | HCIA | Computing",
            font=('Arial', 10)
        )
        self.question_info.grid(row=0, column=0)
        
        # Área da questão - armazenar referência
        self.question_frame = ttk.LabelFrame(main_frame, text="QUESTION", padding="10")
        self.question_frame.grid(row=1, column=0, columnspan=2, sticky=(tk.W, tk.E, tk.N, tk.S), pady=(0, 10))
        self.question_frame.columnconfigure(0, weight=1)
        
        # Enunciado
        self.question_text = tk.Text(
            self.question_frame, 
            height=6, 
            wrap=tk.WORD, 
            font=('Arial', 11),
            bg='white'
        )
        self.question_text.grid(row=0, column=0, sticky=(tk.W, tk.E), pady=(0, 10))
        
        # Scrollbar para o enunciado
        scrollbar = ttk.Scrollbar(self.question_frame, orient=tk.VERTICAL, command=self.question_text.yview)
        scrollbar.grid(row=0, column=1, sticky=(tk.N, tk.S))
        self.question_text.configure(yscrollcommand=scrollbar.set)
        
        # Alternativas - armazenar referência
        self.alternatives_frame = ttk.LabelFrame(main_frame, text="ALTERNATIVES", padding="10")
        self.alternatives_frame.grid(row=2, column=0, columnspan=2, sticky=(tk.W, tk.E), pady=(0, 10))
        self.alternatives_frame.columnconfigure(0, weight=1)
        
        self.alternative_widgets = []
        
        for i in range(4):  # Apenas 4 alternativas (A, B, C, D)
            frame = ttk.Frame(self.alternatives_frame)
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
        self.answer_btn = ttk.Button(
            main_frame, 
            text="Show Answer", 
            command=self.show_answer
        )
        self.answer_btn.grid(row=3, column=0, columnspan=2, pady=(10, 5))
        
        # Área de resposta - armazenar referência
        self.answer_frame = ttk.LabelFrame(main_frame, text="ANSWER AND EXPLANATION", padding="10")
        self.answer_frame.grid(row=4, column=0, columnspan=2, sticky=(tk.W, tk.E, tk.N, tk.S))
        self.answer_frame.columnconfigure(0, weight=1)
        self.answer_frame.rowconfigure(0, weight=1)
        
        self.answer_text = tk.Text(
            self.answer_frame, 
            height=8, 
            wrap=tk.WORD, 
            font=('Arial', 10),
            bg='#f8f8f8'
        )
        self.answer_text.grid(row=0, column=0, sticky=(tk.W, tk.E, tk.N, tk.S))
        
        # Scrollbar para resposta
        answer_scrollbar = ttk.Scrollbar(self.answer_frame, orient=tk.VERTICAL, command=self.answer_text.yview)
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
        if self.current_language == "portuguese":
            self.question_info.config(
                text=f"Questão {self.current_question_index + 1}/{len(self.questions)} | {question['level']} | {question['track']}"
            )
        else:
            self.question_info.config(
                text=f"Question {self.current_question_index + 1}/{len(self.questions)} | {question['level']} | {question['track']}"
            )
        
        # Exibir enunciado
        self.question_text.config(state=tk.NORMAL)
        self.question_text.delete(1.0, tk.END)
        
        if self.current_language == "portuguese":
            # Usar texto traduzido se disponível
            if (self.current_question_index in self.translated_questions and 
                'text_pt' in self.translated_questions[self.current_question_index]):
                self.question_text.insert(1.0, self.translated_questions[self.current_question_index]['text_pt'])
            else:
                # Se não está traduzida, usar o original em inglês
                self.question_text.insert(1.0, question['text_en'])
        else:
            self.question_text.insert(1.0, question['text_en'])
            
        self.question_text.config(state=tk.DISABLED)
        
        # Exibir alternativas
        for i, letra in enumerate(['A', 'B', 'C', 'D']):
            text_widget = self.alternative_widgets[i]
            text_widget.config(state=tk.NORMAL)
            text_widget.delete(1.0, tk.END)
            
            if letra in question['alternatives']:
                text_widget.insert(1.0, question['alternatives'][letra]['text'])
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
            if self.current_language == "portuguese":
                self.answer_text.insert(1.0, "Esta questão não possui gabarito definido.")
            else:
                self.answer_text.insert(1.0, "This question has no defined answer.")
            self.answer_text.config(state=tk.DISABLED)
            return
        
        # Construir o texto da resposta baseado no idioma
        resposta_texto = ""
        
        if self.current_language == "portuguese":
            resposta_texto = f"RESPOSTA CORRETA: {question['correct_answer']}\n\n"
            
            # Explicação em português (original)
            if question['explanation_pt']:
                resposta_texto += f"EXPLICAÇÃO:\n{question['explanation_pt']}\n\n"
        else:
            resposta_texto = f"CORRECT ANSWER: {question['correct_answer']}\n\n"
            
            # Explicação em inglês (traduzida)
            if question['explanation_pt']:
                # Usar explicação traduzida se disponível
                if (self.current_question_index in self.translated_questions and 
                    'explanation_en' in self.translated_questions[self.current_question_index]):
                    resposta_texto += f"EXPLANATION:\n{self.translated_questions[self.current_question_index]['explanation_en']}\n\n"
                else:
                    # Se não está traduzida, traduzir agora
                    self.translate_explanation_to_english()
                    if (self.current_question_index in self.translated_questions and 
                        'explanation_en' in self.translated_questions[self.current_question_index]):
                        resposta_texto += f"EXPLANATION:\n{self.translated_questions[self.current_question_index]['explanation_en']}\n\n"
                    else:
                        # Fallback: usar original em português se a tradução falhar
                        resposta_texto += f"EXPLANATION:\n{question['explanation_pt']}\n\n"
        
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
            
            # Atualizar textos da interface para português
            self.root.title("Questões de Estudo - HCIA Computing")
            self.question_frame.config(text="QUESTÃO")
            self.alternatives_frame.config(text="ALTERNATIVAS")
            self.answer_frame.config(text="RESPOSTA E EXPLICAÇÃO")
            self.answer_btn.config(text="Mostrar Resposta")
            self.prev_btn.config(text="← Anterior")
            self.next_btn.config(text="Próximo →")
            
            # Traduzir a questão atual para português
            self.translate_current_question_to_portuguese()
            
        else:
            self.current_language = "english"
            self.language_btn.config(text="Português")
            
            # Atualizar textos da interface para inglês
            self.root.title("Study Questions - HCIA Computing")
            self.question_frame.config(text="QUESTION")
            self.alternatives_frame.config(text="ALTERNATIVES")
            self.answer_frame.config(text="ANSWER AND EXPLANATION")
            self.answer_btn.config(text="Show Answer")
            self.prev_btn.config(text="← Previous")
            self.next_btn.config(text="Next →")
        
        # Atualizar a questão atual
        self.show_question()
    
    def translate_current_question_to_portuguese(self):
        """Traduz a questão atual para português"""
        try:
            question = self.questions[self.current_question_index]
            
            # Configurar tradutor para português
            self.translator.target = 'pt'
            
            # Traduzir enunciado (do inglês para português)
            translated_text = self.translator.translate(question['text_en'])
            
            # Inicializar entrada no cache se não existir
            if self.current_question_index not in self.translated_questions:
                self.translated_questions[self.current_question_index] = {}
            
            # Armazenar no cache
            self.translated_questions[self.current_question_index]['text_pt'] = translated_text
            
        except Exception as e:
            print(f"Erro ao traduzir questão: {str(e)}")
        finally:
            # Restaurar tradutor para inglês
            self.translator.target = 'en'
    
    def translate_explanation_to_english(self):
        """Traduz a explicação da questão atual para inglês"""
        try:
            question = self.questions[self.current_question_index]
            
            if question['explanation_pt']:
                # Traduzir explicação (do português para inglês)
                translated_explanation = self.translator.translate(question['explanation_pt'])
                
                # Inicializar entrada no cache se não existir
                if self.current_question_index not in self.translated_questions:
                    self.translated_questions[self.current_question_index] = {}
                
                # Armazenar no cache
                self.translated_questions[self.current_question_index]['explanation_en'] = translated_explanation
            
        except Exception as e:
            print(f"Erro ao traduzir explicação: {str(e)}")

def main():
    root = tk.Tk()
    app = StudyApp(root)
    root.mainloop()

if __name__ == "__main__":
    main()
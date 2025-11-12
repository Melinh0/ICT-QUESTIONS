#!/usr/bin/env python3
# -*- coding: utf-8 -*-

import json
import tkinter as tk
from tkinter import ttk, messagebox, simpledialog
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
        self.user_answers = {}  # Armazenar respostas do usuário
        self.performance = {'correct': 0, 'incorrect': 0}  # Estatísticas de desempenho
        
        # Carregar questões
        self.load_questions()
        
        # Criar interface
        self.create_widgets()
        self.show_question()
    
    def load_questions(self):
        """Carrega as questões do arquivo JSON"""
        try:
            json_path = f"questoes_explicadas/explicacoes_completas.json"
            
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
                    
                    # Marcar como correta - APENAS UMA pode ser correta
                    if item['fields']['is_correct']:
                        if questions_dict[question_id]['correct_answer'] is None:
                            questions_dict[question_id]['correct_answer'] = letter
                        else:
                            # Se já existe uma correta, mostrar aviso
                            print(f"AVISO: Questão {question_id} tem múltiplas alternativas corretas. Usando a primeira encontrada: {questions_dict[question_id]['correct_answer']}")
        
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
        
        self.prev_btn = ttk.Button(nav_frame, text="← Previous", command=self.previous_question)
        self.prev_btn.grid(row=0, column=0, padx=(0, 5))
        
        self.next_btn = ttk.Button(nav_frame, text="Next →", command=self.next_question)
        self.next_btn.grid(row=0, column=1, padx=(5, 0))
        
        # Busca
        search_frame = ttk.Frame(controls_frame)
        search_frame.grid(row=0, column=2, padx=(20, 0))
        
        ttk.Label(search_frame, text="Search:").grid(row=0, column=0, padx=(0, 5))
        self.search_var = tk.StringVar()
        self.search_entry = ttk.Entry(search_frame, textvariable=self.search_var, width=20)
        self.search_entry.grid(row=0, column=1, padx=(0, 5))
        self.search_entry.bind('<Return>', lambda e: self.search_questions())
        
        self.search_btn = ttk.Button(search_frame, text="Go", command=self.search_questions)
        self.search_btn.grid(row=0, column=2)
        
        # Informações da questão
        info_frame = ttk.Frame(controls_frame)
        info_frame.grid(row=0, column=3, padx=(20, 0))
        
        self.question_info = ttk.Label(
            info_frame, 
            text="Question 1/150 | HCIA | Computing",
            font=('Arial', 10)
        )
        self.question_info.grid(row=0, column=0)
        
        # Estatísticas de desempenho
        self.performance_label = ttk.Label(
            controls_frame,
            text="Correct: 0 | Incorrect: 0 | Score: 0%",
            font=('Arial', 9)
        )
        self.performance_label.grid(row=0, column=4, padx=(20, 0))
        
        # Área da questão
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
        
        # Alternativas
        self.alternatives_frame = ttk.LabelFrame(main_frame, text="ALTERNATIVES", padding="10")
        self.alternatives_frame.grid(row=2, column=0, columnspan=2, sticky=(tk.W, tk.E), pady=(0, 10))
        self.alternatives_frame.columnconfigure(0, weight=1)
        
        self.alternative_widgets = []
        self.alternative_vars = []  # Variáveis para os RadioButtons
        
        # Variável para grupo de RadioButtons (garante que apenas um seja selecionado)
        self.selected_alternative = tk.StringVar()
        
        for i in range(4):  # Apenas 4 alternativas (A, B, C, D)
            frame = ttk.Frame(self.alternatives_frame)
            frame.grid(row=i, column=0, sticky=(tk.W, tk.E), pady=2)
            frame.columnconfigure(1, weight=1)
            
            # RadioButton para seleção - TODOS usam a mesma variável
            radio = ttk.Radiobutton(
                frame, 
                variable=self.selected_alternative, 
                value=chr(65+i),
                command=lambda idx=i: self.select_alternative(chr(65+idx))
            )
            radio.grid(row=0, column=0, padx=(0, 5))
            
            # Label com a letra da alternativa
            label = ttk.Label(frame, text=f"{chr(65+i)})", font=('Arial', 10, 'bold'))
            label.grid(row=0, column=1, padx=(0, 10), sticky=tk.W)
            
            text_widget = tk.Text(
                frame, 
                height=2, 
                wrap=tk.WORD, 
                font=('Arial', 10),
                bg='white',
                width=60
            )
            text_widget.grid(row=0, column=2, sticky=(tk.W, tk.E))
            
            self.alternative_widgets.append(text_widget)
        
        # Frame para botões de ação
        action_frame = ttk.Frame(main_frame)
        action_frame.grid(row=3, column=0, columnspan=2, pady=(10, 5))
        
        # Botão de ver resposta
        self.answer_btn = ttk.Button(
            action_frame, 
            text="Show Answer", 
            command=self.show_answer
        )
        self.answer_btn.grid(row=0, column=0, padx=(0, 10))
        
        # Botão para editar gabarito
        self.edit_answer_btn = ttk.Button(
            action_frame,
            text="Edit Correct Answer",
            command=self.edit_correct_answer
        )
        self.edit_answer_btn.grid(row=0, column=1, padx=(0, 10))
        
        # Botão para editar explicação
        self.edit_explanation_btn = ttk.Button(
            action_frame,
            text="Edit Explanation",
            command=self.edit_explanation
        )
        self.edit_explanation_btn.grid(row=0, column=2)
        
        # Botão para limpar seleção
        self.clear_btn = ttk.Button(
            action_frame,
            text="Clear Selection",
            command=self.clear_selection
        )
        self.clear_btn.grid(row=0, column=3, padx=(10, 0))
        
        # Área de resposta
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
    
    def search_questions(self):
        """Busca questões pelo texto do enunciado (case-insensitive)"""
        search_term = self.search_var.get().strip().lower()
        
        if not search_term:
            if self.current_language == "portuguese":
                messagebox.showinfo("Busca", "Digite um termo para buscar.")
            else:
                messagebox.showinfo("Search", "Please enter a search term.")
            return
        
        found_indices = []
        
        for i, question in enumerate(self.questions):
            # Buscar no texto original em inglês
            if search_term in question['text_en'].lower():
                found_indices.append(i)
                continue
            
            # Buscar no texto traduzido se disponível
            if (i in self.translated_questions and 
                'text_pt' in self.translated_questions[i] and
                search_term in self.translated_questions[i]['text_pt'].lower()):
                found_indices.append(i)
                continue
            
            # Buscar nas alternativas (inglês)
            for alt in question['alternatives'].values():
                if search_term in alt['text'].lower():
                    found_indices.append(i)
                    break
        
        if not found_indices:
            if self.current_language == "portuguese":
                messagebox.showinfo("Busca", f"Nenhuma questão encontrada com: '{search_term}'")
            else:
                messagebox.showinfo("Search", f"No questions found with: '{search_term}'")
            return
        
        # Se encontrou apenas uma questão, ir para ela
        if len(found_indices) == 1:
            self.current_question_index = found_indices[0]
            self.show_question()
            if self.current_language == "portuguese":
                messagebox.showinfo("Busca", f"Questão {found_indices[0] + 1} encontrada.")
            else:
                messagebox.showinfo("Search", f"Question {found_indices[0] + 1} found.")
        else:
            # Se encontrou múltiplas, mostrar diálogo de seleção
            self.show_search_results_dialog(found_indices, search_term)
    
    def show_search_results_dialog(self, indices, search_term):
        """Mostra diálogo com resultados da busca para seleção"""
        results_window = tk.Toplevel(self.root)
        results_window.title("Search Results")
        results_window.geometry("600x400")
        results_window.transient(self.root)
        results_window.grab_set()
        
        # Frame principal
        main_frame = ttk.Frame(results_window, padding="10")
        main_frame.pack(fill=tk.BOTH, expand=True)
        
        # Título
        if self.current_language == "portuguese":
            title_text = f"Resultados da busca por: '{search_term}' - {len(indices)} questões encontradas"
        else:
            title_text = f"Search results for: '{search_term}' - {len(indices)} questions found"
        
        ttk.Label(main_frame, text=title_text, font=('Arial', 11, 'bold')).pack(anchor=tk.W, pady=(0, 10))
        
        # Frame para lista de resultados
        tree_frame = ttk.Frame(main_frame)
        tree_frame.pack(fill=tk.BOTH, expand=True)
        
        # Treeview para mostrar resultados
        columns = ('id', 'question_preview')
        tree = ttk.Treeview(tree_frame, columns=columns, show='headings', height=15)
        
        # Configurar colunas
        if self.current_language == "portuguese":
            tree.heading('id', text='ID')
            tree.heading('question_preview', text='Prévia da Questão')
        else:
            tree.heading('id', text='ID')
            tree.heading('question_preview', text='Question Preview')
        
        tree.column('id', width=50)
        tree.column('question_preview', width=500)
        
        # Scrollbar para a treeview
        scrollbar = ttk.Scrollbar(tree_frame, orient=tk.VERTICAL, command=tree.yview)
        tree.configure(yscrollcommand=scrollbar.set)
        
        tree.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
        scrollbar.pack(side=tk.RIGHT, fill=tk.Y)
        
        # Preencher com resultados
        for idx in indices:
            question = self.questions[idx]
            
            # Obter prévia do texto da questão
            if self.current_language == "portuguese" and idx in self.translated_questions and 'text_pt' in self.translated_questions[idx]:
                preview_text = self.translated_questions[idx]['text_pt']
            else:
                preview_text = question['text_en']
            
            # Limitar tamanho da prévia
            if len(preview_text) > 100:
                preview_text = preview_text[:100] + "..."
            
            tree.insert('', tk.END, values=(idx + 1, preview_text), tags=(str(idx),))
        
        # Botão para ir para questão selecionada
        def go_to_selected():
            selection = tree.selection()
            if not selection:
                if self.current_language == "portuguese":
                    messagebox.showwarning("Seleção", "Selecione uma questão da lista.")
                else:
                    messagebox.showwarning("Selection", "Please select a question from the list.")
                return
            
            selected_item = selection[0]
            question_idx = int(tree.item(selected_item, 'tags')[0])
            self.current_question_index = question_idx
            results_window.destroy()
            self.show_question()
        
        button_frame = ttk.Frame(main_frame)
        button_frame.pack(fill=tk.X, pady=(10, 0))
        
        if self.current_language == "portuguese":
            ttk.Button(button_frame, text="Ir para Questão Selecionada", command=go_to_selected).pack(side=tk.RIGHT)
            ttk.Button(button_frame, text="Cancelar", command=results_window.destroy).pack(side=tk.RIGHT, padx=(0, 10))
        else:
            ttk.Button(button_frame, text="Go to Selected Question", command=go_to_selected).pack(side=tk.RIGHT)
            ttk.Button(button_frame, text="Cancel", command=results_window.destroy).pack(side=tk.RIGHT, padx=(0, 10))
    
    def select_alternative(self, alternative):
        """Marca uma alternativa como resposta do usuário"""
        question_id = self.current_question_index
        self.user_answers[question_id] = alternative
        
        # Verificar se a resposta está correta
        current_question = self.questions[question_id]
        is_correct = (alternative == current_question['correct_answer'])
        
        # Atualizar estatísticas se ainda não foi verificada esta questão
        if 'checked' not in current_question:
            if is_correct:
                self.performance['correct'] += 1
            else:
                self.performance['incorrect'] += 1
            current_question['checked'] = True
            current_question['was_correct'] = is_correct
        
        self.update_performance_display()
    
    def clear_selection(self):
        """Limpa a seleção atual da questão"""
        self.selected_alternative.set('')
        question_id = self.current_question_index
        
        # Remover resposta do usuário se existir
        if question_id in self.user_answers:
            # Reverter estatísticas se a questão já foi verificada
            current_question = self.questions[question_id]
            if 'checked' in current_question and current_question['checked']:
                if current_question['was_correct']:
                    self.performance['correct'] -= 1
                else:
                    self.performance['incorrect'] -= 1
                
                # Marcar como não verificada
                current_question['checked'] = False
                del current_question['was_correct']
            
            # Remover resposta
            del self.user_answers[question_id]
        
        self.update_performance_display()
        
        if self.current_language == "portuguese":
            messagebox.showinfo("Sucesso", "Seleção limpa!")
        else:
            messagebox.showinfo("Success", "Selection cleared!")
    
    def update_performance_display(self):
        """Atualiza o display de estatísticas de desempenho"""
        total = self.performance['correct'] + self.performance['incorrect']
        if total > 0:
            score = (self.performance['correct'] / total) * 100
            self.performance_label.config(
                text=f"Correct: {self.performance['correct']} | Incorrect: {self.performance['incorrect']} | Score: {score:.1f}%"
            )
        else:
            self.performance_label.config(
                text=f"Correct: {self.performance['correct']} | Incorrect: {self.performance['incorrect']} | Score: 0%"
            )
    
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
        
        # Restaurar seleção do usuário se existir
        if self.current_question_index in self.user_answers:
            self.selected_alternative.set(self.user_answers[self.current_question_index])
        else:
            self.selected_alternative.set('')
        
        # Limpar área de resposta
        self.answer_text.config(state=tk.NORMAL)
        self.answer_text.delete(1.0, tk.END)
        self.answer_text.config(state=tk.DISABLED)
    
    def show_answer(self):
        """Mostra a resposta correta e explicação"""
        if not self.questions:
            return
        
        question = self.questions[self.current_question_index]
        
        if not question['has_answer'] or question['correct_answer'] is None:
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
        
        # Verificar se o usuário respondeu e se acertou
        user_answer = self.user_answers.get(self.current_question_index)
        is_correct = (user_answer == question['correct_answer'])
        
        if self.current_language == "portuguese":
            resposta_texto = f"RESPOSTA CORRETA: {question['correct_answer']}\n\n"
            
            if user_answer:
                if is_correct:
                    resposta_texto += f"SUA RESPOSTA: {user_answer} ✓ CORRETA\n\n"
                else:
                    resposta_texto += f"SUA RESPOSTA: {user_answer} ✗ ERRADA\n\n"
            else:
                resposta_texto += "Você ainda não respondeu esta questão.\n\n"
            
            # Explicação em português (original)
            if question['explanation_pt']:
                resposta_texto += f"EXPLICAÇÃO:\n{question['explanation_pt']}\n\n"
        else:
            resposta_texto = f"CORRECT ANSWER: {question['correct_answer']}\n\n"
            
            if user_answer:
                if is_correct:
                    resposta_texto += f"YOUR ANSWER: {user_answer} ✓ CORRECT\n\n"
                else:
                    resposta_texto += f"YOUR ANSWER: {user_answer} ✗ INCORRECT\n\n"
            else:
                resposta_texto += "You haven't answered this question yet.\n\n"
            
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
    
    def edit_correct_answer(self):
        """Permite editar a resposta correta da questão"""
        if not self.questions:
            return
        
        question = self.questions[self.current_question_index]
        
        # Diálogo para selecionar nova resposta correta
        new_answer = simpledialog.askstring(
            "Edit Correct Answer",
            f"Current correct answer: {question['correct_answer']}\n\nEnter new correct answer (A, B, C, or D):",
            initialvalue=question['correct_answer'] or "A"
        )
        
        if new_answer and new_answer.upper() in ['A', 'B', 'C', 'D']:
            old_answer = question['correct_answer']
            question['correct_answer'] = new_answer.upper()
            question['has_answer'] = True
            
            # Atualizar estatísticas se o usuário já respondeu
            if self.current_question_index in self.user_answers:
                user_answer = self.user_answers[self.current_question_index]
                was_correct = (user_answer == old_answer) if old_answer else False
                is_correct_now = (user_answer == new_answer.upper())
                
                # Ajustar estatísticas se necessário
                if 'checked' in question and question['checked']:
                    if was_correct and not is_correct_now:
                        self.performance['correct'] -= 1
                        self.performance['incorrect'] += 1
                    elif not was_correct and is_correct_now:
                        self.performance['correct'] += 1
                        self.performance['incorrect'] -= 1
                    
                    question['was_correct'] = is_correct_now
            
            self.update_performance_display()
            messagebox.showinfo("Success", f"Correct answer updated to: {new_answer.upper()}")
            self.show_question()  # Atualizar display
        elif new_answer:  # Usuário digitou algo inválido
            messagebox.showerror("Error", "Please enter a valid answer (A, B, C, or D)")
    
    def edit_explanation(self):
        """Permite editar a explicação da questão"""
        if not self.questions:
            return
        
        question = self.questions[self.current_question_index]
        
        # Criar janela de edição
        edit_window = tk.Toplevel(self.root)
        edit_window.title("Edit Explanation")
        edit_window.geometry("600x400")
        edit_window.transient(self.root)
        edit_window.grab_set()
        
        # Frame principal
        main_frame = ttk.Frame(edit_window, padding="10")
        main_frame.pack(fill=tk.BOTH, expand=True)
        
        # Texto da explicação atual
        ttk.Label(main_frame, text="Current Explanation:").pack(anchor=tk.W)
        
        explanation_text = tk.Text(
            main_frame, 
            height=15, 
            wrap=tk.WORD,
            font=('Arial', 10)
        )
        explanation_text.pack(fill=tk.BOTH, expand=True, pady=(5, 10))
        explanation_text.insert(1.0, question['explanation_pt'] or "")
        
        # Frame para botões
        button_frame = ttk.Frame(main_frame)
        button_frame.pack(fill=tk.X)
        
        def save_explanation():
            new_explanation = explanation_text.get(1.0, tk.END).strip()
            question['explanation_pt'] = new_explanation
            
            # Limpar cache de tradução
            if self.current_question_index in self.translated_questions:
                if 'explanation_en' in self.translated_questions[self.current_question_index]:
                    del self.translated_questions[self.current_question_index]['explanation_en']
            
            messagebox.showinfo("Success", "Explanation updated successfully")
            edit_window.destroy()
            
            # Atualizar display se estiver mostrando a resposta
            if self.answer_text.get(1.0, tk.END).strip():
                self.show_answer()
        
        def cancel_edit():
            edit_window.destroy()
        
        ttk.Button(button_frame, text="Save", command=save_explanation).pack(side=tk.RIGHT, padx=(5, 0))
        ttk.Button(button_frame, text="Cancel", command=cancel_edit).pack(side=tk.RIGHT)
    
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
            self.edit_answer_btn.config(text="Editar Gabarito")
            self.edit_explanation_btn.config(text="Editar Explicação")
            self.clear_btn.config(text="Limpar Seleção")
            self.search_btn.config(text="Buscar")
            
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
            self.edit_answer_btn.config(text="Edit Correct Answer")
            self.edit_explanation_btn.config(text="Edit Explanation")
            self.clear_btn.config(text="Clear Selection")
            self.search_btn.config(text="Search")
        
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
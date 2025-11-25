import tkinter as tk
from tkinter import ttk, messagebox, filedialog
import pandas as pd
import random
import hashlib
import time
import sys
import os
import copy
from datetime import datetime
import threading

# =============================================================================
# LÓGICA DE NEGÓCIO E UTILITÁRIOS
# =============================================================================

def sanitizar_nome_arquivo(texto):
    t = str(texto).upper().strip()
    t = t.replace(" ", "_").replace("/", "-").replace("\\", "-").replace(":", "")
    t = t.replace("º", "o").replace("°", "o").replace("ª", "a")
    t = t.replace("Á", "A").replace("É", "E").replace("Í", "I").replace("Ó", "O").replace("Ú", "U")
    t = t.replace("Ã", "A").replace("Õ", "O").replace("Ç", "C")
    return t

def normalizar_nome_cota(texto):
    texto = str(texto).upper().strip()
    if "PPI1" in texto: return "PPI1"
    if "PPI2" in texto: return "PPI2"
    if "RF" in texto: return "RF"
    if "PCD" in texto or "DEFICIÊNCIA" in texto: return "PCD"
    if "AC" in texto or "AMPLA" in texto: return "AC"
    return texto 

def normalizar_nome_turma(texto):
    t = str(texto).upper().strip()
    if "TURMA 1" in t: return "TURMA 1"
    if "TURMA 2" in t: return "TURMA 2"
    if "TURMA 3" in t: return "TURMA 3"
    if "TURMA 4" in t: return "TURMA 4"
    
    if "1º ANO DO ENSINO FUNDAMENTAL" in t or "1° ANO" in t: return "1º ANO EF"
    if "2º ANO DO ENSINO FUNDAMENTAL" in t or "2° ANO" in t: return "2º ANO EF"
    if "3º ANO DO ENSINO FUNDAMENTAL" in t or "3° ANO" in t: return "3º ANO EF"
    if "4º ANO DO ENSINO FUNDAMENTAL" in t or "4° ANO" in t: return "4º ANO EF"
    if "5º ANO DO ENSINO FUNDAMENTAL" in t or "5° ANO" in t: return "5º ANO EF"
    if "6º ANO DO ENSINO FUNDAMENTAL" in t or "6° ANO" in t: return "6º ANO EF"
    if "7º ANO DO ENSINO FUNDAMENTAL" in t or "7° ANO" in t: return "7º ANO EF"
    if "8º ANO DO ENSINO FUNDAMENTAL" in t or "8° ANO" in t: return "8º ANO EF"
    if "9º ANO DO ENSINO FUNDAMENTAL" in t or "9° ANO" in t: return "9º ANO EF"
    
    if "1ª SÉRIE" in t or "1ª SERIE" in t: return "1ª SÉRIE EM"
    if "2ª SÉRIE" in t or "2ª SERIE" in t: return "2ª SÉRIE EM"
    if "3ª SÉRIE" in t or "3ª SERIE" in t: return "3ª SÉRIE EM"
    return t 

# =============================================================================
# INTERFACE GRÁFICA (Frontend)
# =============================================================================

class SorteioApp:
    def __init__(self, root):
        self.root = root
        self.root.title("Sorteio Coluni-UFF v4.4")
        self.root.geometry("1366x768")
        try:
            self.root.state('zoomed') 
        except:
            pass
        
        self.style = ttk.Style()
        self.style.theme_use('clam')
        self.style.configure("Treeview", font=("Arial", 11), rowheight=25)
        self.style.configure("Treeview.Heading", font=("Arial", 11, "bold"), background="#e1e1e1")
        
        self.cor_bg = "#f0f0f0"
        self.cor_primary = "#003366" 
        self.cor_secondary = "#cc0000" 
        self.cor_text = "#333333"
        
        self.root.configure(bg=self.cor_bg)

        # Estado
        self.df = None
        self.hash_arquivo = ""
        self.cancelados_removidos = 0
        self.semente = tk.StringVar()
        self.pasta_resultados = ""
        
        self.turmas_ordenadas = []
        self.indice_turma_atual = 0
        self.ultima_config = None 
        
        self.log_geral = [] 
        self.ORDEM_COTAS = ['PPI1', 'PPI2', 'RF', 'PCD', 'AC']
        self.COTA_FINAL = 'AC'
        
        # Fila de Eventos
        self.fila_de_eventos = []
        self.dados_planilha_atual = []
        self.log_turma_atual = []
        self.turma_atual_nome = ""
        
        self.criar_layout_base()
        self.mostrar_tela_inicial()

    def criar_layout_base(self):
        self.header_frame = tk.Frame(self.root, bg=self.cor_primary, height=80)
        self.header_frame.pack(fill=tk.X, side=tk.TOP)
        self.header_frame.pack_propagate(False)
        
        lbl_titulo = tk.Label(self.header_frame, text="SORTEIO PÚBLICO - COLUNI UFF", 
                              font=("Arial", 24, "bold"), bg=self.cor_primary, fg="white")
        lbl_titulo.pack(pady=20)

        self.footer_frame = tk.Frame(self.root, bg="#cccccc", height=30)
        self.footer_frame.pack(fill=tk.X, side=tk.BOTTOM)
        
        texto_rodape = "Desenvolvido por Alessandro Santos - GitHub: @alessandrosanntos | Versão 4.4"
        lbl_creditos = tk.Label(self.footer_frame, text=texto_rodape,
                                font=("Arial", 10, "bold"), bg="#cccccc", fg="#333333")
        lbl_creditos.pack(pady=5)

        self.main_container = tk.Frame(self.root, bg=self.cor_bg)
        self.main_container.pack(fill=tk.BOTH, expand=True, padx=20, pady=20)

    def limpar_container(self):
        for widget in self.main_container.winfo_children():
            widget.destroy()
            
    def resolver_caminho_arquivo(self, nome_arquivo):
        if getattr(sys, 'frozen', False):
            diretorio_base = os.path.dirname(sys.executable)
        else:
            diretorio_base = os.path.dirname(os.path.abspath(__file__))
        return os.path.join(diretorio_base, nome_arquivo)

    def abrir_documento(self, nome_arquivo_base):
        extensoes = ["", ".txt", ".md"]
        caminho_final = None
        for ext in extensoes:
            tentativa = self.resolver_caminho_arquivo(nome_arquivo_base + ext)
            if os.path.exists(tentativa):
                caminho_final = tentativa
                break
        if caminho_final:
            try:
                os.startfile(caminho_final)
            except Exception as e:
                messagebox.showerror("Erro", f"Erro ao abrir arquivo:\n{e}")
        else:
            messagebox.showerror("Erro", f"Arquivo '{nome_arquivo_base}' não encontrado.")

    # ================= TELA 1: CARREGAMENTO =================
    def mostrar_tela_inicial(self):
        self.limpar_container()
        frame = tk.Frame(self.main_container, bg="white", bd=2, relief=tk.RIDGE)
        frame.place(relx=0.5, rely=0.5, anchor=tk.CENTER, width=600, height=500)

        tk.Label(frame, text="Configuração Inicial", font=("Arial", 18, "bold"), bg="white").pack(pady=20)

        tk.Button(frame, text="Carregar Planilha (base.xlsx)", command=self.carregar_arquivo, 
                  font=("Arial", 12), bg="#e0e0e0", width=30).pack(pady=10)
        
        self.lbl_status = tk.Label(frame, text="Nenhum arquivo carregado", bg="white", fg="red")
        self.lbl_status.pack()

        tk.Label(frame, text="Semente de Auditoria:", font=("Arial", 12), bg="white").pack(pady=(20, 5))
        tk.Entry(frame, textvariable=self.semente, font=("Arial", 14), width=30, justify='center').pack(pady=5)
        tk.Label(frame, text="(Use data/hora, nº concurso, etc.)", font=("Arial", 9), bg="white", fg="gray").pack()

        tk.Button(frame, text="INICIAR SORTEIO", command=self.iniciar_processo, 
                  font=("Arial", 14, "bold"), bg=self.cor_primary, fg="white", width=20).pack(pady=30)
        
        fr_docs = tk.Frame(frame, bg="white")
        fr_docs.pack(side=tk.BOTTOM, pady=20)
        tk.Button(fr_docs, text="1 - Licença", command=lambda: self.abrir_documento("README"), width=15).pack(side=tk.LEFT, padx=5)
        tk.Button(fr_docs, text="2 - Manual", command=lambda: self.abrir_documento("TUTORIAL_USUARIO"), width=15).pack(side=tk.LEFT, padx=5)

    def carregar_arquivo(self):
        path = filedialog.askopenfilename(filetypes=[("Excel", "*.xlsx")])
        if not path: return

        try:
            df = pd.read_excel(path, header=0)
            if len(df.columns) < 4:
                messagebox.showerror("Erro", "A planilha deve ter pelo menos 4 colunas (A a D).")
                return

            df = df.iloc[:, 0:4]
            self.col_insc = 'INSCRICAO'
            self.col_nome = 'NOME'
            self.col_turma = 'TURMA'
            self.col_cota = 'COTA'
            
            df.columns = [self.col_insc, self.col_nome, self.col_turma, self.col_cota]

            for col in df.columns:
                df[col] = df[col].astype(str).str.strip()

            df = df[df[self.col_nome] != 'nan']
            df = df[df[self.col_nome] != '']
            
            total_antes = len(df)
            df = df[df[self.col_nome].str.upper() != 'NÚMERO CANCELADO']
            self.cancelados_removidos = total_antes - len(df)

            df[self.col_turma] = df[self.col_turma].apply(normalizar_nome_turma)
            df[self.col_cota] = df[self.col_cota].apply(normalizar_nome_cota)

            with open(path, 'rb') as f:
                self.hash_arquivo = hashlib.sha256(f.read()).hexdigest()

            self.df = df
            self.lbl_status.config(text=f"Sucesso! {len(df)} inscritos. ({self.cancelados_removidos} cancelados)", fg="green")
            
        except Exception as e:
            messagebox.showerror("Erro", f"Falha ao ler arquivo:\n{e}")

    def iniciar_processo(self):
        if self.df is None or not self.semente.get().strip():
            messagebox.showwarning("Atenção", "Carregue o arquivo e digite a semente.")
            return

        random.seed(self.semente.get())
        semente_segura = sanitizar_nome_arquivo(self.semente.get())
        self.pasta_resultados = f"RESULTADO DO SORTEIO - {semente_segura}"
        if not os.path.exists(self.pasta_resultados):
            os.makedirs(self.pasta_resultados)

        ORDEM = [
            'TURMA 1', 'TURMA 2', 'TURMA 3', 'TURMA 4',
            '1º ANO EF', '2º ANO EF', '3º ANO EF', '4º ANO EF', '5º ANO EF',
            '6º ANO EF', '7º ANO EF', '8º ANO EF', '9º ANO EF',
            '1ª SÉRIE EM', '2ª SÉRIE EM', '3ª SÉRIE EM'
        ]
        existentes = self.df[self.col_turma].unique()
        self.turmas_ordenadas = [t for t in ORDEM if t in existentes]
        for t in existentes:
            if t not in self.turmas_ordenadas: self.turmas_ordenadas.append(t)

        if not self.turmas_ordenadas:
            messagebox.showerror("Erro", "Nenhuma turma encontrada.")
            return

        ts = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
        self.log_geral = [
            "========================================================================",
            "                       REGISTRO DE AUDITORIA",
            "             SORTEIO PÚBLICO PARA INGRESSO NO COLUNI-UFF",
            "========================================================================",
            f"Versão: 4.4 (GUI)",
            f"Início: {ts}",
            f"Semente: {self.semente.get()}",
            f"Hash Base: {self.hash_arquivo}",
            f"Cancelados Removidos: {self.cancelados_removidos}",
            "-" * 72 + "\n"
        ]

        self.indice_turma_atual = 0
        self.mostrar_tela_configuracao_turma()

    # ================= TELA 2: CONFIGURAÇÃO DE VAGAS =================
    def mostrar_tela_configuracao_turma(self):
        self.limpar_container()
        
        if self.indice_turma_atual >= len(self.turmas_ordenadas):
            self.finalizar_sorteio()
            return

        turma_atual = self.turmas_ordenadas[self.indice_turma_atual]
        inscritos = len(self.df[self.df[self.col_turma] == turma_atual])
        
        frame = tk.Frame(self.main_container, bg="white", padx=20, pady=20)
        frame.pack(expand=True, fill=tk.BOTH)

        tk.Label(frame, text=f"Configurando: {turma_atual}", font=("Arial", 22, "bold"), bg="white", fg=self.cor_primary).pack()
        tk.Label(frame, text=f"Inscritos nesta turma: {inscritos}", font=("Arial", 12), bg="white").pack(pady=5)

        if self.ultima_config:
            tk.Button(frame, text="📋 Copiar Vagas da Turma Anterior", 
                      command=self.copiar_anterior, bg="#fffccc").pack(pady=10)

        self.entries = {} 
        grid = tk.Frame(frame, bg="white")
        grid.pack(pady=10)

        tk.Label(grid, text="Cota", font=("Arial", 11, "bold"), bg="white", width=10).grid(row=0, column=0)
        tk.Label(grid, text="Vagas Imediatas", font=("Arial", 11, "bold"), bg="white", width=15).grid(row=0, column=1)
        tk.Label(grid, text="Cadastro Reserva", font=("Arial", 11, "bold"), bg="white", width=15).grid(row=0, column=2)

        for i, cota in enumerate(self.ORDEM_COTAS):
            tk.Label(grid, text=cota, font=("Arial", 12, "bold"), bg="white").grid(row=i+1, column=0, pady=5)
            e_im = tk.Entry(grid, justify='center', font=("Arial", 12))
            e_im.insert(0, "0")
            e_im.grid(row=i+1, column=1, padx=10)
            e_cr = tk.Entry(grid, justify='center', font=("Arial", 12))
            e_cr.insert(0, "0")
            e_cr.grid(row=i+1, column=2, padx=10)
            self.entries[cota] = {'im': e_im, 'cr': e_cr}

        tk.Button(frame, text="CONFIRMAR E IR PARA SORTEIO", 
                  command=lambda: self.validar_e_preparar(turma_atual),
                  font=("Arial", 14, "bold"), bg="#28a745", fg="white", width=30).pack(pady=20)
        
        tk.Button(frame, text="PULAR TURMA (Sem Vagas)", 
                  command=lambda: self.acao_pular_turma_config(turma_atual),
                  font=("Arial", 10, "bold"), bg="#dc3545", fg="white", width=30).pack(pady=10)

    def copiar_anterior(self):
        if not self.ultima_config: return
        for cota, vals in self.ultima_config.items():
            self.entries[cota]['im'].delete(0, tk.END)
            self.entries[cota]['im'].insert(0, str(vals['im']))
            self.entries[cota]['cr'].delete(0, tk.END)
            self.entries[cota]['cr'].insert(0, str(vals['cr']))

    def acao_pular_turma_config(self, turma):
        if messagebox.askyesno("Confirmar Pulo", f"Tem certeza que deseja pular a turma {turma}?\n\nIsso registrará no log que não houve sorteio para esta turma."):
            ts = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
            log_pulado = [
                "========================================================================",
                f"RESULTADO: {turma}",
                f"Data/Hora: {ts}",
                f"Semente: {self.semente.get()}",
                "-" * 72,
                "*** TURMA PULADA PELO OPERADOR (SEM VAGAS OFERTADAS) ***",
                "*** NENHUM SORTEIO REALIZADO ***"
            ]
            
            timestamp_arq = datetime.now().strftime("%Y%m%d_%H%M%S")
            nome_limpo = sanitizar_nome_arquivo(turma)
            nome_arq = f"Resultado_{nome_limpo}_{timestamp_arq}.txt"
            caminho = os.path.join(self.pasta_resultados, nome_arq)
            
            try:
                with open(caminho, "w", encoding="utf-8") as f:
                    f.write("\n".join(log_pulado))
                self.log_geral.extend(log_pulado)
                self.log_geral.append("\n" + "="*72 + "\n")
            except Exception as e:
                print(f"Erro ao salvar log pulado: {e}")
            
            self.ir_para_proxima()

    def validar_e_preparar(self, turma):
        config = {}
        try:
            for cota, inputs in self.entries.items():
                im = int(inputs['im'].get())
                cr = int(inputs['cr'].get())
                if im < 0 or cr < 0: raise ValueError
                config[cota] = {'im': im, 'cr': cr}
        except ValueError:
            messagebox.showerror("Erro", "Digite apenas números inteiros positivos.")
            return

        msg_confirmacao = f"Confirma o quadro de vagas para {turma}?\n\n"
        for cota in self.ORDEM_COTAS:
            msg_confirmacao += f"{cota}: {config[cota]['im']} (IM) / {config[cota]['cr']} (CR)\n"
        
        if not messagebox.askyesno("Confirmar Configuração", msg_confirmacao):
            return

        self.ultima_config = config
        self.preparar_fila_de_eventos(turma, config)
        self.mostrar_tela_execucao(turma)

    # ================= LÓGICA DE PRÉ-CÁLCULO (EVENTOS) =================
    def preparar_fila_de_eventos(self, turma, config):
        self.fila_de_eventos = [] 
        self.dados_planilha_atual = []
        self.log_turma_atual = []
        self.turma_atual_nome = turma
        
        ts = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
        header = [
            "========================================================================",
            "                       REGISTRO DE AUDITORIA",
            "             SORTEIO PÚBLICO PARA INGRESSO NO COLUNI-UFF",
            "========================================================================",
            f"Versão: 4.4",
            f"Turma: {turma}",
            f"Data/Hora: {ts}",
            f"Semente: {self.semente.get()}",
            f"Hash Base: {self.hash_arquivo}",
            "-" * 72
        ]
        self.log_turma_atual.extend(header)
        self.fila_de_eventos.append({'tipo': 'header', 'texto': f"Iniciando sorteio para {turma}...\n"})

        df_turma = self.df[self.df[self.col_turma] == turma].copy()
        
        cotas_processo = ['PPI1', 'PPI2', 'RF', 'PCD']
        cota_ac = 'AC'
        sobras_im = 0
        sobras_cr = 0

        # 1. Cotas Prioritárias
        for cota in cotas_processo:
            lim_im = config[cota]['im']
            lim_cr = config[cota]['cr']
            
            self.fila_de_eventos.append({'tipo': 'titulo_cota', 'texto': f"--- CATEGORIA: {cota} ---"})
            
            if lim_im == 0 and lim_cr == 0:
                self.fila_de_eventos.append({'tipo': 'info', 'texto': f"   0 vagas ofertadas para {cota}."})
                continue

            candidatos = df_turma[df_turma[self.col_cota] == cota].to_dict('records')
            random.shuffle(candidatos) 
            
            # Processa Imediatas
            if lim_im > 0:
                self.fila_de_eventos.append({'tipo': 'subtitulo', 'texto': f"   [VAGAS IMEDIATAS: {lim_im}]"})
                vencedores = candidatos[:lim_im]
                restantes = candidatos[lim_im:]
                
                if not vencedores:
                    self.fila_de_eventos.append({'tipo': 'info', 'texto': f"   Sem inscritos. {lim_im} vaga(s) transferida(s) para AC."})
                    sobras_im += lim_im
                else:
                    for i, p in enumerate(vencedores):
                        self.fila_de_eventos.append({
                            'tipo': 'sorteio', 'dados': p, 'ordem': i+1, 
                            'msg_cota': f"{cota} - IMEDIATA", 'situacao': 'IMEDIATA', 'origem': cota
                        })
                    diff = lim_im - len(vencedores)
                    if diff > 0:
                        self.fila_de_eventos.append({'tipo': 'info', 'texto': f"   {diff} vaga(s) não preenchida(s) transferida(s) para AC."})
                        sobras_im += diff
            else:
                restantes = candidatos 

            # Processa Reserva
            if lim_cr > 0:
                self.fila_de_eventos.append({'tipo': 'subtitulo', 'texto': f"   [CADASTRO RESERVA: {lim_cr}]"})
                vencedores_cr = restantes[:lim_cr]
                if not vencedores_cr:
                    self.fila_de_eventos.append({'tipo': 'info', 'texto': f"   Sem inscritos para reserva. {lim_cr} vaga(s) transferida(s) para AC."})
                    sobras_cr += lim_cr
                else:
                    for i, p in enumerate(vencedores_cr):
                        self.fila_de_eventos.append({
                            'tipo': 'sorteio', 'dados': p, 'ordem': i+1, 
                            'msg_cota': f"{cota} - RESERVA", 'situacao': 'CADASTRO RESERVA', 'origem': cota
                        })
                    diff = lim_cr - len(vencedores_cr)
                    if diff > 0:
                        self.fila_de_eventos.append({'tipo': 'info', 'texto': f"   {diff} vaga(s) reserva não preenchida(s) transferida(s) para AC."})
                        sobras_cr += diff
            
            self.fila_de_eventos.append({'tipo': 'espaco'})

        # 2. Ampla Concorrência
        orig_im_ac = config[cota_ac]['im']
        orig_cr_ac = config[cota_ac]['cr']
        total_im_ac = orig_im_ac + sobras_im
        total_cr_ac = orig_cr_ac + sobras_cr
        
        self.fila_de_eventos.append({'tipo': 'titulo_cota', 'texto': f"--- CATEGORIA: {cota_ac} (AMPLA CONCORRÊNCIA) ---"})
        self.fila_de_eventos.append({'tipo': 'info', 'texto': f"   Total Imediatas: {total_im_ac} ({orig_im_ac} Originais + {sobras_im} Sobras)"})
        self.fila_de_eventos.append({'tipo': 'info', 'texto': f"   Total Reserva: {total_cr_ac} ({orig_cr_ac} Originais + {sobras_cr} Sobras)"})
        
        candidatos_ac = df_turma[df_turma[self.col_cota] == cota_ac].to_dict('records')
        random.shuffle(candidatos_ac)
        
        # Imediatas AC
        vencedores_im_ac = candidatos_ac[:total_im_ac]
        restantes_ac = candidatos_ac[total_im_ac:]
        
        if total_im_ac > 0 and len(vencedores_im_ac) > 0:
            self.fila_de_eventos.append({'tipo': 'subtitulo', 'texto': f"   [RESULTADO IMEDIATAS AC]"})
            for i, p in enumerate(vencedores_im_ac):
                self.fila_de_eventos.append({
                    'tipo': 'sorteio', 'dados': p, 'ordem': i+1, 
                    'msg_cota': f"{cota_ac} - IMEDIATA", 'situacao': 'IMEDIATA', 'origem': cota_ac
                })
        elif total_im_ac > 0 and len(vencedores_im_ac) == 0:
             self.fila_de_eventos.append({'tipo': 'subtitulo', 'texto': f"   [RESULTADO IMEDIATAS AC]"})
             self.fila_de_eventos.append({'tipo': 'info', 'texto': "   Nenhum candidato."})

        # Reserva AC
        vencedores_cr_ac = restantes_ac[:total_cr_ac]
        
        if total_cr_ac > 0:
            self.fila_de_eventos.append({'tipo': 'subtitulo', 'texto': f"\n   [RESULTADO RESERVA AC]"})
            if vencedores_cr_ac:
                for i, p in enumerate(vencedores_cr_ac):
                    self.fila_de_eventos.append({
                        'tipo': 'sorteio', 'dados': p, 'ordem': i+1, 
                        'msg_cota': f"{cota_ac} - RESERVA", 'situacao': 'CADASTRO RESERVA', 'origem': cota_ac
                    })
            else:
                 self.fila_de_eventos.append({'tipo': 'info', 'texto': "   Nenhum candidato."})

        self.fila_de_eventos.append({'tipo': 'fim', 'texto': "\n--- FIM DA TURMA ---"})

    # ================= TELA 3: EXECUÇÃO =================
    def mostrar_tela_execucao(self, turma):
        self.limpar_container()
        
        frame_exec = tk.Frame(self.main_container, bg=self.cor_bg)
        frame_exec.pack(fill=tk.BOTH, expand=True)

        tk.Label(frame_exec, text=f"Sorteando: {turma}", font=("Arial", 20, "bold"), bg=self.cor_bg, fg=self.cor_primary).pack(pady=10)

        self.txt_resultado = tk.Text(frame_exec, font=("Courier New", 14), width=100, height=20)
        scrollbar = tk.Scrollbar(frame_exec, command=self.txt_resultado.yview)
        self.txt_resultado.config(yscrollcommand=scrollbar.set)
        self.txt_resultado.pack(side=tk.LEFT, fill=tk.BOTH, expand=True, padx=10)
        scrollbar.pack(side=tk.RIGHT, fill=tk.Y)

        frame_rodape = tk.Frame(frame_exec, bg=self.cor_bg)
        frame_rodape.pack(side=tk.BOTTOM, fill=tk.X, padx=10, pady=10)

        self.btn_acao = tk.Button(frame_rodape, text="SORTEAR PRÓXIMO >>", 
                                  command=self.executar_proximo_passo,
                                  font=("Arial", 14, "bold"), bg="#007bff", fg="white", height=2)
        self.btn_acao.pack(fill=tk.X)

        self.executar_proximo_passo()

    def executar_proximo_passo(self):
        if not self.fila_de_eventos:
            self.finalizar_turma_atual()
            return

        evento = self.fila_de_eventos.pop(0)
        tipo = evento['tipo']

        if tipo in ['header', 'titulo_cota', 'subtitulo', 'info', 'espaco']:
            texto = evento.get('texto', "")
            self.txt_resultado.insert(tk.END, texto + "\n")
            self.txt_resultado.see(tk.END)
            if tipo not in ['espaco', 'subtitulo']:
                self.log_turma_atual.append(texto)
            self.root.after(50, self.executar_proximo_passo)

        elif tipo == 'sorteio':
            self.btn_acao.config(state=tk.DISABLED, text="Sorteando...")
            threading.Thread(target=self.animar_e_revelar, args=(evento,)).start()
            
        elif tipo == 'fim':
            self.txt_resultado.insert(tk.END, evento['texto'] + "\n")
            self.log_turma_atual.append(evento['texto'])
            self.btn_acao.config(text="VER PLANILHA E FINALIZAR", bg="#28a745", command=self.finalizar_turma_atual)

    def animar_e_revelar(self, evento):
        # 1. Animação Sorteando (5s)
        tempo_anim = 5.0 
        fim = time.time() + tempo_anim
        simbolos = ['|', '/', '-', '\\']
        i = 0
        
        self.txt_resultado.insert(tk.END, "   🎲 Sorteando... ")
        self.txt_resultado.see(tk.END)
        
        while time.time() < fim:
            char = simbolos[i % 4]
            self.txt_resultado.insert(tk.END, char)
            self.txt_resultado.see(tk.END)
            time.sleep(0.1)
            self.txt_resultado.delete("end-1c")
            i += 1
        self.txt_resultado.delete("end-19c", "end")

        # 2. Exibe Inscrição
        inscricao = evento['dados'][self.col_insc]
        self.txt_resultado.insert(tk.END, f"   Inscrição: {inscricao}")
        self.txt_resultado.see(tk.END)
        
        # 3. Delay e Mensagem de Busca
        time.sleep(2)
        msg_temp = " - Nome: Aguarde. Buscando informações..."
        self.txt_resultado.insert(tk.END, msg_temp)
        self.txt_resultado.see(tk.END)
        time.sleep(3)
        self.txt_resultado.delete(f"end-{len(msg_temp)}c", tk.END)
        
        # 4. Completa com o Nome
        nome = evento['dados'][self.col_nome]
        self.txt_resultado.insert(tk.END, f" - Nome: {nome}\n")
        self.txt_resultado.see(tk.END)
        
        # 5. Registra
        texto_log = f"   Classificação {evento['ordem']}: {nome} (Insc: {inscricao})"
        self.log_turma_atual.append(f"{texto_log} [{evento['msg_cota']}]")
        
        self.dados_planilha_atual.append({
            'Classificação': evento['ordem'],
            'Nome': nome,
            'Inscrição': inscricao,
            'Cota': evento['origem'],
            'Situação': evento['situacao']
        })

        self.root.after(0, lambda: self.btn_acao.config(state=tk.NORMAL, text="SORTEAR PRÓXIMO >>"))

    def finalizar_turma_atual(self):
        timestamp_arq = datetime.now().strftime("%Y%m%d_%H%M%S")
        nome_limpo = sanitizar_nome_arquivo(self.turma_atual_nome)
        nome_arq = f"Resultado_{nome_limpo}_{timestamp_arq}.txt"
        
        caminho = os.path.join(self.pasta_resultados, nome_arq)
        try:
            with open(caminho, "w", encoding="utf-8") as f:
                f.write("\n".join(self.log_turma_atual))
            self.log_geral.extend(self.log_turma_atual)
            self.log_geral.append("\n" + "="*72 + "\n")
        except Exception as e:
            print(f"Erro ao salvar log: {e}")

        self.exibir_planilha_final(self.turma_atual_nome)

    def exibir_planilha_final(self, turma):
        popup = tk.Toplevel(self.root)
        popup.title(f"Resultado - {turma}")
        popup.geometry("1000x600")
        popup.transient(self.root)
        popup.grab_set()
        
        cols = ("Classificação", "Nome", "Inscrição", "Cota", "Situação")
        tree = ttk.Treeview(popup, columns=cols, show='headings')
        
        tree.column("Classificação", width=50, anchor=tk.CENTER); tree.heading("Classificação", text="Class.")
        tree.column("Nome", width=400, anchor=tk.W); tree.heading("Nome", text="Nome")
        tree.column("Inscrição", width=100, anchor=tk.CENTER); tree.heading("Inscrição", text="Inscrição")
        tree.column("Cota", width=80, anchor=tk.CENTER); tree.heading("Cota", text="Cota")
        tree.column("Situação", width=150, anchor=tk.CENTER); tree.heading("Situação", text="Situação")
        
        for item in self.dados_planilha_atual:
            tag = 'reserva' if item['Situação'] == 'CADASTRO RESERVA' else 'imediata'
            tree.insert("", tk.END, values=(item['Classificação'], item['Nome'], item['Inscrição'], item['Cota'], item['Situação']), tags=(tag,))

        tree.tag_configure('imediata', background='#e8f5e9')
        tree.tag_configure('reserva', background='#fff3e0')
        
        tree.pack(fill=tk.BOTH, expand=True, padx=10, pady=10)
        
        def fechar():
            popup.destroy()
            self.ir_para_proxima()
            
        tk.Button(popup, text="Fechar e Próxima Turma", command=fechar, bg=self.cor_primary, fg="white", font=("Arial", 12)).pack(pady=10)
        self.root.wait_window(popup)

    def ir_para_proxima(self):
        self.indice_turma_atual += 1
        self.mostrar_tela_configuracao_turma()

    def finalizar_sorteio(self):
        self.limpar_container()
        caminho = os.path.join(self.pasta_resultados, "RESUMO_GERAL_CONSOLIDADO.txt")
        with open(caminho, "w", encoding="utf-8") as f:
            f.write("\n".join(self.log_geral))
            
        lbl = tk.Label(self.main_container, text="Sorteio Finalizado!", font=("Arial", 24), fg="green", bg=self.cor_bg)
        lbl.pack(expand=True)
        
        fr_fim = tk.Frame(self.main_container, bg=self.cor_bg)
        fr_fim.pack(pady=20)
        
        # ATUALIZAÇÃO 2: Texto e Tamanho do Botão
        tk.Button(fr_fim, text="Resultado do Sorteio (Abrir Pasta de Registros)", 
                  command=lambda: os.startfile(self.pasta_resultados),
                  font=("Arial", 14), width=40).pack(pady=10)
        
        tk.Button(fr_fim, text="Fechar e Encerrar", command=self.root.quit,
                  font=("Arial", 14, "bold"), bg="#cc0000", fg="white", width=40).pack(pady=10)

if __name__ == "__main__":
    root = tk.Tk()
    app = SorteioApp(root)
    root.mainloop()
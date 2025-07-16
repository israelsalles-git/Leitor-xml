import os
import xml.etree.ElementTree as ET
import pandas as pd
import time
from watchdog.observers import Observer
from watchdog.events import FileSystemEventHandler
import customtkinter as ctk
from tkinter import filedialog, messagebox
import threading
from collections import OrderedDict
import traceback

# Configuração do tema da interface
ctk.set_appearance_mode("System")
ctk.set_default_color_theme("blue")

class XMLHandler(FileSystemEventHandler):
    def __init__(self, app):
        self.app = app
    
    def on_created(self, event):
        if not event.is_directory and event.src_path.lower().endswith('.xml'):
            self.app.process_new_xml(event.src_path)

class XMLProcessorApp:
    def __init__(self):
        self.window = ctk.CTk()
        self.window.title("Monitor de XML para Excel - Controle de Notas Fiscais")
        self.window.geometry("900x650")
        
        # Variáveis de controle
        self.monitoring = False
        self.observer = None
        self.output_file = "relatorio_notas.xlsx"
        self.data = OrderedDict()
        
        # Inicializar a barra de status primeiro
        self.status_bar = ctk.CTkLabel(
            self.window,
            text="Pronto",
            anchor="w",
            font=("Arial", 10)
        )
        self.status_bar.pack(fill="x", padx=20, pady=(0, 10))
        
        # Configuração da interface
        self.create_widgets()
        
    def create_widgets(self):
        # Frame principal
        self.main_frame = ctk.CTkFrame(self.window)
        self.main_frame.pack(pady=20, padx=20, fill="both", expand=True)
        
        # Título
        self.title_label = ctk.CTkLabel(
            self.main_frame, 
            text="Sistema de Monitoramento de Notas Fiscais",
            font=("Arial", 20, "bold")
        )
        self.title_label.pack(pady=(10, 20))
        
        # Frame de configurações
        self.settings_frame = ctk.CTkFrame(self.main_frame)
        self.settings_frame.pack(fill="x", padx=10, pady=10)
        
        # Seletor de pasta
        self.folder_frame = ctk.CTkFrame(self.settings_frame)
        self.folder_frame.pack(fill="x", pady=5)
        
        ctk.CTkLabel(
            self.folder_frame, 
            text="Pasta para monitorar:",
            font=("Arial", 12)
        ).pack(anchor="w")
        
        self.folder_path = ctk.StringVar()
        self.folder_entry = ctk.CTkEntry(
            self.folder_frame, 
            textvariable=self.folder_path,
            width=500
        )
        self.folder_entry.pack(side="left", fill="x", expand=True, padx=(0, 10))
        
        ctk.CTkButton(
            self.folder_frame,
            text="Procurar",
            command=self.select_folder
        ).pack(side="left")
        
        # Seletor de arquivo de saída
        self.output_frame = ctk.CTkFrame(self.settings_frame)
        self.output_frame.pack(fill="x", pady=5)
        
        ctk.CTkLabel(
            self.output_frame, 
            text="Arquivo Excel de saída:",
            font=("Arial", 12)
        ).pack(anchor="w")
        
        self.output_path = ctk.StringVar(value=self.output_file)
        self.output_entry = ctk.CTkEntry(
            self.output_frame, 
            textvariable=self.output_path,
            width=500
        )
        self.output_entry.pack(side="left", fill="x", expand=True, padx=(0, 10))
        
        ctk.CTkButton(
            self.output_frame,
            text="Procurar",
            command=self.select_output_file
        ).pack(side="left")
        
        # Frame de controle
        self.control_frame = ctk.CTkFrame(self.main_frame)
        self.control_frame.pack(pady=15)
        
        # Botões de ação
        self.start_button = ctk.CTkButton(
            self.control_frame,
            text="▶ Iniciar Monitoramento",
            command=self.start_monitoring,
            fg_color="green",
            hover_color="dark green",
            width=200
        )
        self.start_button.pack(side="left", padx=10)
        
        self.stop_button = ctk.CTkButton(
            self.control_frame,
            text="■ Parar Monitoramento",
            command=self.stop_monitoring,
            fg_color="red",
            hover_color="dark red",
            state="disabled",
            width=200
        )
        self.stop_button.pack(side="left", padx=10)
        
        self.process_button = ctk.CTkButton(
            self.control_frame,
            text="↻ Processar XMLs Existentes",
            command=self.process_existing_files,
            width=200
        )
        self.process_button.pack(side="left", padx=10)
        
        # Área de log
        self.log_frame = ctk.CTkFrame(self.main_frame)
        self.log_frame.pack(fill="both", expand=True, padx=10, pady=(0, 10))
        
        ctk.CTkLabel(
            self.log_frame, 
            text="Log de Atividades:",
            font=("Arial", 12)
        ).pack(anchor="w", padx=5, pady=(5, 0))
        
        self.log_text = ctk.CTkTextbox(
            self.log_frame,
            font=("Consolas", 10),
            wrap="word"
        )
        self.log_text.pack(fill="both", expand=True, padx=5, pady=5)
        self.log_message("Sistema inicializado. Aguardando configurações...")
    
    def select_folder(self):
        folder_selected = filedialog.askdirectory()
        if folder_selected:
            self.folder_path.set(folder_selected)
            self.log_message(f"Pasta selecionada: {folder_selected}")
    
    def select_output_file(self):
        file_selected = filedialog.asksaveasfilename(
            defaultextension=".xlsx",
            filetypes=[("Excel files", "*.xlsx"), ("All files", "*.*")],
            initialfile="relatorio_notas.xlsx"
        )
        if file_selected:
            self.output_path.set(file_selected)
            self.output_file = file_selected
            self.log_message(f"Arquivo de saída definido: {file_selected}")
    
    def log_message(self, message):
        timestamp = time.strftime('%Y-%m-%d %H:%M:%S')
        self.log_text.configure(state="normal")
        self.log_text.insert("end", f"[{timestamp}] {message}\n")
        self.log_text.see("end")
        self.log_text.configure(state="disabled")
        self.status_bar.configure(text=message)
    
    def start_monitoring(self):
        folder = self.folder_path.get()
        if not folder:
            messagebox.showerror("Erro", "Selecione uma pasta para monitorar!")
            return
        
        if not os.path.isdir(folder):
            messagebox.showerror("Erro", "A pasta selecionada não existe!")
            return
        
        self.output_file = self.output_path.get()
        if not self.output_file:
            messagebox.showerror("Erro", "Defina um arquivo de saída!")
            return
        
        self.monitoring = True
        self.start_button.configure(state="disabled")
        self.stop_button.configure(state="normal")
        self.process_button.configure(state="normal")
        
        self.log_message(f"Iniciando monitoramento da pasta: {folder}")
        self.log_message(f"Arquivo de saída: {self.output_file}")
        
        # Processar arquivos existentes
        self.process_existing_files()
        
        # Iniciar monitoramento em thread separada
        self.monitor_thread = threading.Thread(
            target=self.run_monitor,
            args=(folder,),
            daemon=True
        )
        self.monitor_thread.start()
    
    def run_monitor(self, folder):
        event_handler = XMLHandler(self)
        self.observer = Observer()
        self.observer.schedule(event_handler, folder, recursive=True)
        self.observer.start()
        
        try:
            while self.monitoring:
                time.sleep(1)
        except Exception as e:
            self.log_message(f"Erro no monitoramento: {str(e)}")
            self.observer.stop()
        
        self.observer.join()
    
    def stop_monitoring(self):
        self.monitoring = False
        if self.observer:
            self.observer.stop()
            self.observer.join()
        
        self.start_button.configure(state="normal")
        self.stop_button.configure(state="disabled")
        self.log_message("Monitoramento parado.")
    
    def process_existing_files(self):
        folder = self.folder_path.get()
        if not folder or not os.path.isdir(folder):
            self.log_message("Erro: Pasta não selecionada ou inválida")
            return
        
        self.log_message("Processando arquivos XML existentes na pasta...")
        
        xml_files = [f for f in os.listdir(folder) if f.lower().endswith('.xml')]
        total = len(xml_files)
        processed = 0
        errors = 0
        
        for i, xml_file in enumerate(xml_files, 1):
            file_path = os.path.join(folder, xml_file)
            if self.process_xml(file_path):
                processed += 1
            else:
                errors += 1
            
            # Atualizar status a cada 10 arquivos
            if i % 10 == 0 or i == total:
                self.log_message(f"Progresso: {i}/{total} arquivos processados")
        
        self.log_message(f"Processamento concluído: {processed} sucessos, {errors} erros")
        if processed > 0:
            self.save_to_excel()
    
    def process_new_xml(self, file_path):
        if self.process_xml(file_path):
            self.save_to_excel()
    
    def process_xml(self, file_path):
        try:
            tree = ET.parse(file_path)
            root = tree.getroot()
            
            # Namespace padrão da NF-e
            ns = {'nfe': 'http://www.portalfiscal.inf.br/nfe'}
            
            # Encontrar a tag infNFe que contém os dados principais
            infNFe = root.find('.//nfe:infNFe', ns)
            if infNFe is None:
                self.log_message(f"Arquivo {os.path.basename(file_path)} não é uma NF-e válida")
                return False
            
            # Extrair número da nota (chave única)
            nNF = infNFe.find('.//nfe:nNF', ns)
            numero_nota = nNF.text if nNF is not None else None
            
            if not numero_nota:
                self.log_message(f"Arquivo {os.path.basename(file_path)} não contém número da nota")
                return False
            
            # Verificar se a nota já existe
            if numero_nota in self.data:
                self.log_message(f"Nota {numero_nota} já existe - atualizando dados")
            
            # Extrair data de emissão
            dhEmi = infNFe.find('.//nfe:dhEmi', ns)
            data_emissao = dhEmi.text[:10] if dhEmi is not None else None
            
            # Extrair dados do destinatário (cliente)
            dest = infNFe.find('.//nfe:dest', ns)
            if dest is not None:
                xNome = dest.find('.//nfe:xNome', ns)
                nome_cliente = xNome.text if xNome is not None else "CONSUMIDOR NÃO IDENTIFICADO"
            else:
                nome_cliente = "CONSUMIDOR NÃO IDENTIFICADO"
            
            # Extrair valor total da nota
            vNF = infNFe.find('.//nfe:vNF', ns)
            valor_total = float(vNF.text) if vNF is not None else 0.0
            
            # Adicionar ou atualizar os dados da nota
            self.data[numero_nota] = {
                'Número da Nota': numero_nota,
                'Data Emissão': data_emissao,
                'Nome do Cliente': nome_cliente,
                'Valor Total': valor_total,
                'Arquivo Origem': os.path.basename(file_path),
                'Data Processamento': time.strftime('%Y-%m-%d %H:%M:%S'),
                'Status': 'ATUALIZADA' if numero_nota in self.data else 'NOVA'
            }
            
            return True
            
        except ET.ParseError:
            self.log_message(f"Erro: Arquivo {os.path.basename(file_path)} não é um XML válido")
            return False
        except Exception as e:
            self.log_message(f"Erro ao processar {os.path.basename(file_path)}: {str(e)}")
            return False
    
    def save_to_excel(self):
        if not self.data:
            self.log_message("Nenhum dado para salvar no Excel")
            return
        
        try:
            # Converter OrderedDict para DataFrame
            df = pd.DataFrame(list(self.data.values()))
            
            # Ordenar por data de emissão (mais recente primeiro)
            df = df.sort_values(by='Data Emissão', ascending=False)
            
            # Verificar se o arquivo já existe
            if os.path.exists(self.output_file):
                try:
                    # Carregar dados existentes
                    existing_df = pd.read_excel(self.output_file, engine='openpyxl')
                    
                    # Criar dicionário com dados existentes
                    existing_data = {row['Número da Nota']: row for _, row in existing_df.iterrows()}
                    
                    # Atualizar com os novos dados (sobrescrevendo se existirem)
                    existing_data.update(self.data)
                    
                    # Criar novo DataFrame consolidado
                    df = pd.DataFrame(list(existing_data.values()))
                except Exception as e:
                    self.log_message(f"Aviso: Não foi possível ler arquivo existente - criando novo")
            
            # Salvar para Excel
            df.to_excel(
                self.output_file,
                index=False,
                engine='openpyxl',
                sheet_name='Notas Fiscais'
            )
            
            self.log_message(f"Relatório salvo com {len(df)} notas (incluindo atualizações)")
            
        except PermissionError:
            self.log_message("Erro: Não foi possível salvar - arquivo está aberto ou sem permissão")
        except Exception as e:
            self.log_message(f"Erro crítico ao salvar Excel: {str(e)}")
            self.log_message(traceback.format_exc())
    
    def run(self):
        self.window.mainloop()

if __name__ == "__main__":
    app = XMLProcessorApp()
    app.run()

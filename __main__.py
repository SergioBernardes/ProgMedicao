import os
import tkinter as tk
from tkinter import ttk, messagebox, filedialog
from tkinter.font import Font
import pandas as pd
from datetime import datetime
from matplotlib import pyplot as plt
from matplotlib.backends.backend_tkagg import FigureCanvasTkAgg, NavigationToolbar2Tk
from matplotlib.ticker import FuncFormatter
import numpy as np
import webbrowser

# Definição dos meses centralizada
MESES = {
    '01': 'Janeiro', '02': 'Fevereiro', '03': 'Março', '04': 'Abril', '05': 'Maio', '06': 'Junho',
    '07': 'Julho', '08': 'Agosto', '09': 'Setembro', '10': 'Outubro', '11': 'Novembro', '12': 'Dezembro'
}

class AutocompleteCombobox(ttk.Combobox):
    def __init__(self, *args, **kwargs):
        super().__init__(*args, **kwargs)
        self._completion_list = []
        self._hits = []
        self._hit_index = 0
        self.position = 0
        self.bind('<KeyRelease>', self.handle_keyrelease)
        self['values'] = []

    def set_completion_list(self, completion_list):
        self._completion_list = sorted(str(item) for item in completion_list)
        self._hits = []
        self._hit_index = 0
        self.position = 0
        self['values'] = self._completion_list

    def autocomplete(self, delta=0):
        if delta:
            self.delete(self.position, tk.END)
        else:
            self.position = len(self.get())

        _hits = [item for item in self._completion_list if item.lower().startswith(self.get().lower())]
        self._hits = _hits

        if _hits:
            self._hit_index = (self._hit_index + delta) % len(_hits)
            self.delete(0, tk.END)
            self.insert(0, self._hits[self._hit_index])
            self.select_range(self.position, tk.END)

    def handle_keyrelease(self, event):
        if event.keysym in ('BackSpace', 'Left', 'Right', 'Up', 'Down'):
            return
        self.autocomplete()

class DataFrameViewer(tk.Tk):
    def __init__(self, dataframe):
        super().__init__()
        self.title("Programa Medição")
        self.state('zoomed')
        self.dataframe_cleaned = self.clean_dataframe(dataframe)
        self.verification_dataframe = self.create_verification_dataframe()
        self.notebook = ttk.Notebook(self)
        self.notebook.pack(fill="both", expand=True)
        self.table_frames = {}
        self.create_tabs()
        self.update_last_update()
        self.graph_index = 0
        self.add_author_label()

    def generate_status_matrix(self):
        """Gera a matriz de status de clientes por mês (Aberto, Fechado, Inativo, Pendente)."""
        # Selecionar colunas relevantes
        df_status = self.dataframe_cleaned[['CLIENTE', 'ABA', 'STATUS', 'Nº MEDIÇÃO']].copy()
        df_status = df_status.dropna(subset=['CLIENTE', 'ABA', 'STATUS', 'Nº MEDIÇÃO'])

        # Gerar lista de clientes únicos
        clientes = df_status['CLIENTE'].unique().tolist()
        meses = [f'24{str(i).zfill(2)}' for i in range(1, 13)]

        # Inicializar a matriz
        status_matrix = []

        for cliente in clientes:
            cliente_row = []
            cliente_data = df_status[df_status['CLIENTE'] == cliente]

            # Obter os 4 primeiros dígitos da primeira ocorrência de 'Nº MEDIÇÃO'
            n_medicao = cliente_data['Nº MEDIÇÃO'].iloc[0].split('-')[0][:4]

            # Adicionar o código de medição e o nome do cliente
            cliente_row.append(n_medicao)  # Nova coluna com os 4 primeiros dígitos
            cliente_row.append(cliente)

            cliente_status = {str(aba): status for aba, status in zip(cliente_data['ABA'], cliente_data['STATUS'])}

            ultimo_status = None
            primeiro_pendente = False  # Indica se o primeiro P já foi encontrado

            for mes in meses:
                if mes in cliente_status:
                    status_atual = cliente_status[mes].strip().upper()

                    if status_atual in ['ATIVO', 'AG. FAT.', 'PARCIAL']:
                        cliente_row.append('A')  # Aberto (Ativo)
                        ultimo_status = 'A'
                    elif status_atual == 'FINALIZADO':
                        cliente_row.append('F')  # Fechado (Finalizado)
                        ultimo_status = 'F'
                    else:
                        cliente_row.append('O')  # Outros status, mapeados como "O"
                        ultimo_status = 'O'
                else:
                    if ultimo_status == 'A' and not primeiro_pendente:
                        cliente_row.append('P')  # Pendente
                        primeiro_pendente = True
                        ultimo_status = 'P'
                    elif ultimo_status == 'F':
                        cliente_row.append('X')  # Inativo após Fechado (Finalizado)
                    elif primeiro_pendente:
                        cliente_row.append('')  # Vazio após o primeiro P
                    else:
                        cliente_row.append('')  # Nenhum status conhecido antes de P

            status_matrix.append(cliente_row)

        return status_matrix
    
    def clean_dataframe(self, df):
        df_cleaned = df.copy()
        df_cleaned['ABA'] = df_cleaned['ABA'].astype(str)  # Garantir que a coluna ABA seja string
        df_cleaned['VALOR FATURADO'] = pd.to_numeric(df_cleaned['VALOR FATURADO'], errors='coerce')
        df_cleaned['GLOSA - MANUTENÇÃO'] = pd.to_numeric(df_cleaned['GLOSA - MANUTENÇÃO'], errors='coerce').fillna(0)
        df_cleaned['DESC COMERCIAL'] = pd.to_numeric(df_cleaned['DESC COMERCIAL'], errors='coerce').fillna(0)
        df_cleaned['KM EXCEDENTE'] = pd.to_numeric(df_cleaned['KM EXCEDENTE'], errors='coerce').fillna(0)
        df_cleaned['MULTA CONTRATUAL'] = pd.to_numeric(df_cleaned['MULTA CONTRATUAL'], errors='coerce').fillna(0)
        df_cleaned['AJUSTES / ACRÉCIMOS'] = pd.to_numeric(df_cleaned['AJUSTES / ACRÉCIMOS'], errors='coerce').fillna(0)
        df_cleaned['QTDE LOCADOS'] = pd.to_numeric(df_cleaned['QTDE LOCADOS'], errors='coerce')
        df_cleaned['QTDE RESERVA'] = pd.to_numeric(df_cleaned['QTDE RESERVA'], errors='coerce')
        return df_cleaned

    def create_verification_dataframe(self):
        epsilon = 0.05  # Margem de erro para desconsiderar variações de até 1 centavo
        df = self.clean_dataframe(self.dataframe_cleaned)
        df['DIF FAT/MED'] = (
            (df['PREVISÃO DE MEDIÇÃO'] - df['GLOSA - MANUTENÇÃO'] - df['DESC COMERCIAL'] +
            df['KM EXCEDENTE'] + df['MULTA CONTRATUAL'] + df['AJUSTES / ACRÉCIMOS']) - df['VALOR FATURADO']
        ).round(2)
        
        # Aplicar a margem de erro para ignorar pequenas diferenças
        df['DIF FAT/MED'] = df['DIF FAT/MED'].apply(lambda x: 0 if abs(x) <= epsilon else x)
        
        verification_df = df[['ABA', 'CLIENTE', 'Nº MEDIÇÃO', 'RESP MEDIÇÃO', 'SITUAÇÃO MED.', 'VALOR FATURADO', 
                            'PREVISÃO DE MEDIÇÃO', 'GLOSA - MANUTENÇÃO', 'DESC COMERCIAL', 'KM EXCEDENTE', 
                            'MULTA CONTRATUAL', 'AJUSTES / ACRÉCIMOS', 'DIF FAT/MED', 'ADM CONTRATO']]
        
        # Filtrar linhas onde DIF FAT/MED é diferente de 0
        verification_df = verification_df[verification_df['DIF FAT/MED'] != 0]
        
        return verification_df
       
    def create_tabs(self):
        self.tabs_info = [
            ("Controle de Medição", self.create_dataframe_viewer),
            ("Gráficos", self.create_graphs_page),
            ("Fechamentos", self.create_closures_page),
            ("Comparar Meses", self.create_comparison_page),
            ("Treinamentos", self.create_trainings_page),
            ("Verificação de Faturamento", self.create_verification_page),
            ("Acompanhamento de Status", self.create_status_tracking_page)  # Nova aba
        ]
       
        for tab_name, tab_method in self.tabs_info:
            frame = ttk.Frame(self.notebook)
            self.notebook.add(frame, text=tab_name)
            tab_method(frame)

    def add_author_label(self):
        author_label = ttk.Label(self, text="Desenvolvido por Sérgio Filho", font=("Helvetica", 10), foreground="gray")
        author_label.pack(side="bottom", pady=5)

    def create_dataframe_viewer(self, page):
        self.create_title(page, "RELATORIO GERAL DE MEDIÇÃO")
        self.update_subtitle(page)
        filter_frame = self.create_filter_frame(page)
        self.create_buttons(filter_frame)
        self.create_filter_buttons(page)
        self.create_treeview(page)

    def create_title(self, page, text):
        title_label = ttk.Label(page, text=text, font=("Helvetica", 20))
        title_label.pack(side="top", pady=10)

    def update_subtitle(self, page):
        self.subtitle_label = ttk.Label(page, text="", font=("Helvetica", 12))
        self.subtitle_label.pack(side="top", pady=5)

    def update_last_update(self):
        now = datetime.now()
        current_time = now.strftime("%d/%m/%Y %H:%M:%S")
        self.subtitle_label.config(text=f"Relatório atualizado às {current_time}")

    def create_filter_frame(self, page):
        filter_frame = ttk.Frame(page)
        filter_frame.pack(side="top", fill="x", padx=10, pady=5)
        columns = self.dataframe_cleaned.columns.tolist()

        def create_filter_row(row, label_text, entry_var, column_select):
            ttk.Label(filter_frame, text=label_text).grid(row=row, column=0, padx=5, pady=5, sticky="w")
            entry = ttk.Entry(filter_frame, textvariable=entry_var, width=40)
            entry.grid(row=row, column=1, padx=5, pady=5)
            ttk.Label(filter_frame, text="NA COLUNA:").grid(row=row, column=2, padx=5, pady=5, sticky="w")
            column_select.set_completion_list(columns)
            column_select.grid(row=row, column=3, padx=5, pady=5)
            if column_select['values']:
                column_select.current(0)

        self.search_var1 = tk.StringVar()
        self.column_select1 = AutocompleteCombobox(filter_frame, width=55)
        create_filter_row(0, "VALOR PROCURADO 1:", self.search_var1, self.column_select1)
        self.search_var2 = tk.StringVar()
        self.column_select2 = AutocompleteCombobox(filter_frame, width=55)
        create_filter_row(1, "VALOR PROCURADO 2:", self.search_var2, self.column_select2)

        self.column_select1.set('')
        self.column_select2.set('')

        return filter_frame

    def create_buttons(self, filter_frame):
        buttons = [
            ("Aplicar Filtro", self.apply_filter, 0, 4),
            ("Limpar Filtro", self.clear_filter, 1, 4),
            ("Exportar Tabela", self.export_table, 0, 5),
            ("Atualizar Relatório", self.update_data, 1, 5)
        ]
        for text, command, row, column in buttons:
            ttk.Button(filter_frame, text=text, command=command).grid(row=row, column=column, padx=5, pady=5)

    def create_filter_buttons(self, page):
        filter_buttons_frame = ttk.Frame(page)
        filter_buttons_frame.pack(side="top", fill="x", padx=10, pady=5)

        ttk.Label(filter_buttons_frame, text="Filtros Rápidos:", font=("Helvetica", 10)).pack(side="left", padx=5)

        months = [f"240{i}" for i in range(1, 10)] + [f"24{i}" for i in range(10, 13)]
        unique_months = self.dataframe_cleaned['ABA'].unique().tolist()

        filter_buttons_frame1 = ttk.Frame(filter_buttons_frame)
        filter_buttons_frame1.pack(side="top", fill="x", padx=5, pady=5)

        filter_buttons_frame2 = ttk.Frame(filter_buttons_frame)
        filter_buttons_frame2.pack(side="top", fill="x", padx=5, pady=5)

        buttons = []
        for month in months:
            if month in unique_months:
                buttons.append((month, lambda m=month: self.apply_quick_filter(m)))

        for idx, (text, command) in enumerate(buttons):
            frame = filter_buttons_frame1 if idx < 16 else filter_buttons_frame2
            button = ttk.Button(frame, text=text, command=command)
            button.pack(side="left", padx=5, pady=5)

    def create_treeview(self, page):
        tree_frame = ttk.Frame(page)
        tree_frame.pack(fill="both", expand=True)

        tree_scroll_y = ttk.Scrollbar(tree_frame, orient="vertical")
        tree_scroll_y.pack(side="right", fill="y")

        tree_scroll_x = ttk.Scrollbar(tree_frame, orient="horizontal")
        tree_scroll_x.pack(side="bottom", fill="x")

        self.treeview = ttk.Treeview(tree_frame, yscrollcommand=tree_scroll_y.set, xscrollcommand=tree_scroll_x.set, show="headings")
        self.treeview.pack(expand=True, fill="both")

        tree_scroll_y.config(command=self.treeview.yview)
        tree_scroll_x.config(command=self.treeview.xview)

        self.populate_treeview(self.dataframe_cleaned)

    def setup_treeview_columns(self, dataframe):
        self.treeview["columns"] = dataframe.columns.tolist()
        default_font = Font()
        for col in self.treeview["columns"]:
            self.treeview.heading(col, text=col, anchor=tk.W, command=lambda c=col: self.sort_column(c))
            column_width = default_font.measure(col.title())
            self.treeview.column(col, width=column_width, stretch=False)

    def populate_treeview(self, dataframe):
        # Limpa todos os itens da Treeview
        self.treeview.delete(*self.treeview.get_children())

        # Define as colunas da Treeview com base no dataframe
        self.treeview["columns"] = dataframe.columns.tolist()
        default_font = Font()
        
        # Configura as colunas e adiciona o evento de ordenação para cada cabeçalho
        for col in self.treeview["columns"]:
            self.treeview.heading(col, text=col, anchor=tk.W, command=lambda c=col: self.sort_column(c))  # Adiciona a funcionalidade de ordenação
            column_width = default_font.measure(col.title())
            self.treeview.column(col, width=column_width, stretch=False)

        # Insere os dados no Treeview
        for _, row in dataframe.iterrows():
            self.treeview.insert("", "end", values=row.tolist())

        # Ajusta as larguras das colunas e reaplica o comando de ordenação
        self.adjust_column_widths()
        self.reapply_sorting()

    def reapply_sorting(self):
        # Reaplica o comando de ordenação para cada coluna do cabeçalho
        for col in self.treeview["columns"]:
            self.treeview.heading(col, command=lambda c=col: self.sort_column(c))

    def sort_column(self, col):
        # Ordena a coluna selecionada
        try:
            data = [(self.treeview.set(child, col), child) for child in self.treeview.get_children("")]
            data.sort(reverse=False)
            for i, (_, child) in enumerate(data):
                self.treeview.move(child, "", i)
        except Exception as e:
            print(f"Erro ao ordenar a coluna {col}: {e}")

    def adjust_column_widths(self):
        for col in self.treeview["columns"]:
            max_width = Font().measure(col.title())
            for row in self.treeview.get_children():
                row_value = self.treeview.item(row, 'values')[self.treeview["columns"].index(col)]
                row_width = Font().measure(str(row_value))
                if row_width > max_width:
                    max_width = row_width
            self.treeview.column(col, width=max_width)

    def sort_column(self, col):
        data = [(self.treeview.set(child, col), child) for child in self.treeview.get_children("")]
        data.sort()
        for i, (_, child) in enumerate(data):
            self.treeview.move(child, "", i)

        for i, item in enumerate(self.treeview.get_children()):
            self.treeview.item(item, tags=("evenrow" if i % 2 == 0 else "oddrow",))

    def apply_filter(self):
        value1 = self.search_var1.get().strip()
        column1 = self.column_select1.get().strip()
        value2 = self.search_var2.get().strip()
        column2 = self.column_select2.get().strip()

        if value1 and column1 not in ('', ' ') and value2 and column2 not in ('', ' '):
            mask1 = self.dataframe_cleaned[column1].astype(str).str.contains(value1, case=False, na=False)
            mask2 = self.dataframe_cleaned[column2].astype(str).str.contains(value2, case=False, na=False)
            filtered_df = self.dataframe_cleaned[mask1 & mask2]
        elif value1 and column1 not in ('', ' '):
            filtered_df = self.dataframe_cleaned[self.dataframe_cleaned[column1].astype(str).str.contains(value1, case=False, na=False)]
        else:
            filtered_df = self.dataframe_cleaned

        self.populate_treeview(filtered_df)

    def apply_quick_filter(self, month):
        month_str = str(month)
        filtered_df = self.dataframe_cleaned[self.dataframe_cleaned['ABA'] == month_str]
        self.populate_treeview(filtered_df)

    def clear_filter(self):
        self.search_var1.set("")
        self.search_var2.set("")
        self.column_select1.set('')
        self.column_select2.set('')
        self.populate_treeview(self.dataframe_cleaned)

    def export_table(self):
        now = datetime.now()
        current_time = now.strftime("%d.%m.%Y_%H-%M-%S")
        file_path = f"C:/Users/{os.getlogin()}/Desktop/RELATORIO_GERAL_MEDICAO_EXPORTADO_{current_time}.xlsx"
        self.dataframe_cleaned.to_excel(file_path, index=False)
        messagebox.showinfo("Exportar Tabela", f"Tabela exportada com sucesso para {file_path}")

    def update_data(self):
        # Recarregar os dados
        self._data()  # Atualiza o dataframe com os novos dados

        # Atualizar a legenda de "última atualização"
        self.update_last_update()

        # Recarregar os gráficos com os novos dados
        self.refresh_graphs()

    def _data(self):
        username = os.getlogin()
        file_paths = [
            f'C:\\Users\\{username}\\EBEC\\NC - Medicao - Documentos\\003 - POWER BI MEDIÇÃO\\001 - RELATORIOS BI\\001 - Fechamento\\RELATORIO GERAL MEDIÇÃO.xlsx',
            f'C:\\Users\\joana.conceicao.EBEC-SA.000\\EBEC\\NC - Medicao - Documentos\\003 - POWER BI MEDIÇÃO\\001 - RELATORIOS BI\\001 - Fechamento\\RELATORIO GERAL MEDIÇÃO.xlsx',
            f'C:\\Users\\{username}\\EBEC\\NC - Medicao - Documentos.000\\003 - POWER BI MEDIÇÃO\\001 - RELATORIOS BI\\001 - Fechamento\\RELATORIO GERAL MEDIÇÃO.xlsx'
        ]

        for file_path in file_paths:
            if os.path.exists(file_path):
                df = pd.read_excel(file_path)
                break
        else:
            file_path = filedialog.askopenfilename(title="Selecione o arquivo RELATORIO GERAL MEDIÇÃO", filetypes=[("Excel files", "*.xlsx")])
            if not file_path:
                messagebox.showwarning("Aviso", "Arquivo não selecionado.")
                return
            df = pd.read_excel(file_path)

        # Atualizar o dataframe limpo com os novos dados
        self.dataframe_cleaned = self.clean_dataframe(df)

        # Atualizar o dataframe de verificação
        self.verification_dataframe = self.create_verification_dataframe()

        # Recarregar a tabela na visualização
        self.populate_treeview(self.dataframe_cleaned)

        # Recarregar os gráficos com os novos dados
        self.refresh_graphs()

    def create_graphs_page(self, page):
        self.graphs = []
        self.graph_frame = ttk.Frame(page)
        self.graph_frame.pack(fill="both", expand=True)

        self.create_graphs()

        self.left_button = ttk.Button(page, text="<", command=self.show_prev_graph)
        self.left_button.pack(side="left", padx=10, pady=10)

        self.right_button = ttk.Button(page, text=">", command=self.show_next_graph)
        self.right_button.pack(side="right", padx=10, pady=10)

        self.show_graph(0)

    def create_graphs(self):
        self.create_graph(self.refresh_graph1, "Valores de Medição / Faturamento")
        self.create_graph(self.refresh_graph9, "Valores Variáveis")  # Novo gráfico adicionado
        self.create_graph(self.refresh_graph10, "Total por CR")  # Novo gráfico adicionado        
        self.create_graph(self.refresh_graph2, "Clientes com Medição Aberta/Fechada por Mês")
        self.create_graph(self.refresh_graph3, "Contagem de RESP MEDIÇÃO por Mês")
        self.create_graph(self.refresh_graph4, "Novos Clientes vs Clientes Finalizados por Mês")
        self.create_graph(self.refresh_graph5, "Total de Carros Locados por Mês")
        self.create_graph(self.refresh_graph6, "Diferença Faturamento/Medição")
        self.create_graph(self.refresh_graph7, "Contagem de Situação por ABA")
        self.create_graph(self.refresh_graph8, "Contagem de Situação por PESSOA")

    def create_graph(self, refresh_method, title):
        figure = plt.Figure(figsize=(12, 6))
        canvas = FigureCanvasTkAgg(figure, master=self.graph_frame)
        ax = figure.add_subplot(111)
        toolbar = NavigationToolbar2Tk(canvas, self.graph_frame)
        toolbar.update()
        canvas.get_tk_widget().pack(side="top", fill="both", expand=True)
        toolbar.pack_forget()  # Esconde a barra de ferramentas inicialmente
        graph = {'canvas': canvas, 'ax': ax, 'title': title, 'toolbar': toolbar}
        self.graphs.append(graph)
        refresh_method(graph)

    def show_graph(self, index):
        for graph in self.graphs:
            graph['canvas'].get_tk_widget().pack_forget()
            graph['toolbar'].pack_forget()

        graph = self.graphs[index]
        graph['canvas'].get_tk_widget().pack(side="top", fill="both", expand=True)
        graph['toolbar'].pack(side="bottom", fill="x")

    def show_next_graph(self):
        self.graph_index = (self.graph_index + 1) % len(self.graphs)
        self.show_graph(self.graph_index)

    def show_prev_graph(self):
        self.graph_index = (self.graph_index - 1) % len(self.graphs)
        self.show_graph(self.graph_index)

    def refresh_graph1(self, graph):
        ax = graph['ax']
        ax.clear()
        aba_values = self.dataframe_cleaned['ABA'].unique().tolist()
        val_fat_values = self.dataframe_cleaned.groupby('ABA')['VALOR FATURADO'].sum().tolist()
        prev_med_values = self.dataframe_cleaned.groupby('ABA')['PREVISÃO DE MEDIÇÃO'].sum().tolist()
        glosa_values = self.dataframe_cleaned.groupby('ABA')['GLOSA - MANUTENÇÃO'].sum().tolist()
        desc_com_values = self.dataframe_cleaned.groupby('ABA')['DESC COMERCIAL'].sum().tolist()
        km_exc_values = self.dataframe_cleaned.groupby('ABA')['KM EXCEDENTE'].sum().tolist()
        multa_values = self.dataframe_cleaned.groupby('ABA')['MULTA CONTRATUAL'].sum().tolist()
        ajustes_values = self.dataframe_cleaned.groupby('ABA')['AJUSTES / ACRÉCIMOS'].sum().tolist()

        # Cálculo da nova linha total
        linha_total = [-glosa - desc + km + multa + ajuste for glosa, desc, km, multa, ajuste in zip(glosa_values, desc_com_values, km_exc_values, multa_values, ajustes_values)]

        x = np.arange(len(aba_values))
        width = 0.35

        # Barras
        bars1 = ax.bar(x - width / 2, val_fat_values, width, label='Valor Faturado', color='#003f70')
        bars2 = ax.bar(x + width / 2, prev_med_values, width, label='Previsão de Medição', color='#00afa0')

        # Rótulos para as barras
        for bar, value in zip(bars1, val_fat_values):
            if value > 0:
                ax.text(bar.get_x() + bar.get_width() / 2, bar.get_height(), f'R$ {value:,.2f}', ha='center', va='bottom',
                        color=bar.get_facecolor(), rotation=90)
        for bar, value in zip(bars2, prev_med_values):
            if value > 0:
                ax.text(bar.get_x() + bar.get_width() / 2, bar.get_height(), f'R$ {value:,.2f}', ha='center', va='bottom',
                        color=bar.get_facecolor(), rotation=90)

        # Eixo principal
        ax.set_ylim(0, 30000000)
        ax.set_ylabel('R$ (Eixo Principal)')

        # Eixo secundário para a linha total
        ax2 = ax.twinx()
        line_total, = ax2.plot(x, linha_total, marker='o', color='#F15A22', label='Variáveis')

        # Ajuste do limite do eixo secundário
        ax2.set_ylim(0, 1500000)
        ax2.set_ylabel('R$ (Eixo Secundário)')

        # Rótulos para a linha total
        for x_value, total, glosa, desc, km, multa, ajuste in zip(x, linha_total, glosa_values, desc_com_values, km_exc_values, multa_values, ajustes_values):
            # Valor total acima do marcador
            if total != 0:
                ax2.text(x_value, total, f'R$ {total:,.2f}', ha='center', va='bottom', color='#F15A22')
            
            # Valores positivos e negativos abaixo do marcador
            sum_negatives = -glosa - desc
            sum_positives = km + multa + ajuste

            # Valor positivo (em verde) e valor negativo (em vermelho)
            if sum_positives != 0:
                ax2.text(x_value, total - (total * 0.05), f'({sum_positives:,.2f})', ha='center', va='top', color='#00FA3C', fontsize=10)
            if sum_negatives != 0:
                ax2.text(x_value, total - (total * 0.15), f'({sum_negatives:,.2f})', ha='center', va='top', color='#FA0000', fontsize=10)

        # Ajuste do posicionamento das linhas para ficarem centralizadas
        ax2.set_xticks(x)
        ax.set_xticks(x)
        ax.set_xticklabels(aba_values)
        ax2.set_xticklabels(aba_values)

        # Legendas combinadas
        bars = [bars1, bars2]
        lines = [line_total]
        labels = [bar.get_label() for bar in bars] + [line.get_label() for line in lines]
        ax.legend(bars + lines, labels, loc='upper right')

        # Formatadores de moeda
        def currency_formatter(x, pos):
            return 'R$ {:,.2f}'.format(x).replace(',', 'x').replace('.', ',').replace('x', '.')

        ax.yaxis.set_major_formatter(FuncFormatter(currency_formatter))
        ax2.yaxis.set_major_formatter(FuncFormatter(currency_formatter))

        graph['canvas'].draw()

    def refresh_graph2(self, graph):
        ax = graph['ax']
        ax.clear()
        aba_values = self.dataframe_cleaned['ABA'].unique().tolist()
        sim_counts = []
        nao_counts = []

        for aba in aba_values:
            aba_df = self.dataframe_cleaned[self.dataframe_cleaned['ABA'] == aba]
            sim_count = aba_df.dropna(subset=['ENVIO FAT', 'FAT MEDIÇÃO']).shape[0]
            nao_count = len(aba_df) - sim_count
            sim_counts.append(sim_count)
            nao_counts.append(nao_count)

        total_counts = [sim + nao for sim, nao in zip(sim_counts, nao_counts)]
        sim_percents = [sim / total * 100 for sim, total in zip(sim_counts, total_counts)]
        nao_percents = [nao / total * 100 for nao, total in zip(nao_counts, total_counts)]

        bar_width = 0.6
        indices = list(range(len(aba_values)))

        p1 = ax.barh(indices, sim_percents, bar_width, label='Finalizados', color='#00afa0')
        p2 = ax.barh(indices, nao_percents, bar_width, left=sim_percents, label='Não Finalizados', color='#a09ba2')

        ax.set_yticks(indices)
        ax.set_yticklabels(aba_values)
        ax.set_title(graph['title'])
        ax.get_xaxis().set_visible(False)

        for rect1, rect2, sim_count, nao_count, total in zip(p1, p2, sim_counts, nao_counts, total_counts):
            ax.text(rect1.get_x() + rect1.get_width() - 1, rect1.get_y() + rect1.get_height() / 2, f'{sim_count}', ha='right', va='center', color='black')
            ax.text(rect2.get_x() + 1, rect2.get_y() + rect2.get_height() / 2, f'{nao_count}', ha='left', va='center', color='black')

        graph['canvas'].draw()

    def refresh_graph3(self, graph):
        ax = graph['ax']
        ax.clear()
        aba_values = self.dataframe_cleaned['ABA'].unique().tolist()
        resp_medicao_values = self.dataframe_cleaned['RESP MEDIÇÃO'].unique().tolist()

        counts = {aba: {resp: 0 for resp in resp_medicao_values} for aba in aba_values}

        for aba in aba_values:
            aba_df = self.dataframe_cleaned[self.dataframe_cleaned['ABA'] == aba]
            for resp in resp_medicao_values:
                counts[aba][resp] = aba_df[aba_df['RESP MEDIÇÃO'] == resp].shape[0]

        aba_indices = range(len(aba_values))
        bar_width = 0.6
        bottom = np.zeros(len(aba_values))

        colors = ['#003f70', '#00afa0', '#ff7f0e', '#d62728', '#9467bd', '#8c564b', '#e377c2']

        for resp, color in zip(resp_medicao_values, colors):
            values = [counts[aba][resp] for aba in aba_values]
            bars = ax.bar(aba_indices, values, bar_width, bottom=bottom, label=resp, color=color)
            bottom += values

            for bar in bars:
                height = bar.get_height()
                if height > 0:
                    ax.text(bar.get_x() + bar.get_width() / 2, bar.get_y() + height / 2, f'{height}', ha='center', va='center', fontsize=8, color='black')

        ax.set_xticks(aba_indices)
        ax.set_xticklabels(aba_values)
        ax.set_title(graph['title'])
        ax.legend(loc='upper right')

        graph['canvas'].draw()

    def refresh_graph4(self, graph):
        ax = graph['ax']
        ax.clear()
        aba_values = self.dataframe_cleaned['ABA'].unique().tolist()

        new_clients_counts = {aba: 0 for aba in aba_values}
        finalized_clients_counts = {aba: 0 for aba in aba_values}

        for aba in aba_values:
            aba_df = self.dataframe_cleaned[self.dataframe_cleaned['ABA'] == aba]
            finalized_clients_counts[aba] = aba_df[aba_df['STATUS'] == 'FINALIZADO']['CLIENTE'].nunique()

        seen_clients = set()
        for idx, row in self.dataframe_cleaned.iterrows():
            if row['CLIENTE'] not in seen_clients:
                seen_clients.add(row['CLIENTE'])
                new_clients_counts[row['ABA']] += 1

        # Limitar novos clientes para 5 em janeiro (2401)
        if '2401' in new_clients_counts:
            new_clients_counts['2401'] = 5

        aba_indices = range(len(aba_values))
        bar_width = 0.6

        new_clients_values = [new_clients_counts[aba] for aba in aba_values]
        finalized_clients_values = [finalized_clients_counts[aba] for aba in aba_values]

        p1 = ax.barh(aba_indices, new_clients_values, bar_width, label='Novos Clientes', color='#00afa0')
        p2 = ax.barh(aba_indices, finalized_clients_values, bar_width, left=new_clients_values, label='Clientes Finalizados', color='#a09ba2')

        ax.set_yticks(aba_indices)
        ax.set_yticklabels(aba_values)
        ax.set_title(graph['title'])
        ax.legend(loc='upper right')

        for bar1, bar2, value1, value2 in zip(p1, p2, new_clients_values, finalized_clients_values):
            width1 = bar1.get_width()
            width2 = bar2.get_width()
            ax.text(width1 / 2, bar1.get_y() + bar1.get_height() / 2, f'{value1}', ha='center', va='center', color='black')
            ax.text(width1 + width2 / 2, bar2.get_y() + bar2.get_height() / 2, f'{value2}', ha='center', va='center', color='black')

        graph['canvas'].draw()

    def refresh_graph5(self, graph):
        ax = graph['ax']
        ax.clear()

        self.dataframe_cleaned['QTDE LOCADOS'] = pd.to_numeric(self.dataframe_cleaned['QTDE LOCADOS'], errors='coerce')
        self.dataframe_cleaned = self.dataframe_cleaned.dropna(subset=['QTDE LOCADOS'])

        valid_resp_medicao = self.dataframe_cleaned['RESP MEDIÇÃO'].dropna().unique().tolist()

        aba_values = self.dataframe_cleaned['ABA'].unique().tolist()
        grouped = self.dataframe_cleaned.groupby(['ABA', 'RESP MEDIÇÃO'])['QTDE LOCADOS'].sum().unstack().fillna(0)

        grouped = grouped[valid_resp_medicao]

        colors = ['#003f70', '#00afa0', '#ff7f0e', '#d62728', '#9467bd', '#8c564b', '#e377c2']
        grouped.plot(kind='bar', stacked=True, ax=ax, color=colors[:len(valid_resp_medicao)])

        ax.set_xticks(range(len(aba_values)))
        ax.set_xticklabels(aba_values)
        ax.set_title(graph['title'])
        ax.legend(title='RESP MEDIÇÃO', bbox_to_anchor=(1.05, 1), loc='upper right')

        for i in range(len(aba_values)):
            for j, (resp, value) in enumerate(grouped.iloc[i].items()):
                if value > 0:
                    ax.text(i, sum(grouped.iloc[i, :j + 1]) - value / 2, f'{int(value)}', ha='center', va='center', color='black')

        graph['canvas'].draw()

    def refresh_graph6(self, graph):
        ax = graph['ax']
        ax.clear()

        aba_counts = self.verification_dataframe['ABA'].value_counts().sort_index()

        aba_counts.plot(kind='bar', ax=ax, color='#003f70')

        ax.set_xlabel('MÊS')
        ax.set_ylabel('Diferenças')
        ax.set_title('Diferença Faturamento/Medição')

        for i, (index, value) in enumerate(aba_counts.items()):
            ax.text(i, value, str(value), ha='center', va='bottom', color='black')

        graph['canvas'].draw()

    def refresh_graph7(self, graph):
        ax = graph['ax']
        ax.clear()

        # Definir o mapeamento de situações para as categorias desejadas
        situation_mapping = {
            'PARCIAL': 'FAT. PARCIAL',
            'AG. CLIENTE': 'AG. CLIENTE/APROV.',
            'AG. APROV.': 'AG. CLIENTE/APROV.',
            'AG. MANUT.': 'AG. MANUT./COMERCIAL',
            'AG. COMERCIAL': 'AG. MANUT./COMERCIAL',
            'AG. FAT.': 'AG. FAT.',
            '1° TENTATIVA': 'ENVIO S/ APROV.',
            '2° TENTATIVA': 'ENVIO S/ APROV.',
            '3° TENTATIVA': 'ENVIO S/ APROV.',
            'ENVIO S/ APROV.': 'ENVIO S/ APROV.',
            'CNPJ': 'CNPJ/LET\'S',
            'LET\'S': 'CNPJ/LET\'S',
            'AG. CANCEL FAT.': 'AG. CANCEL FAT.',
            'AG. DOC': 'AG. DOC',
            # Outras situações vão para "OUTROS"
        }

        # Filtrar e agrupar os dados, ignorando valores vazios
        df_filtered = self.dataframe_cleaned.dropna(subset=['ABA', 'SITUAÇÃO MED.']).copy()  # Usar .copy() para evitar a cópia de visão

        # Aplicar o mapeamento às situações de medição
        df_filtered.loc[:, 'SITUAÇÃO MED.'] = df_filtered['SITUAÇÃO MED.'].apply(
            lambda x: situation_mapping.get(x, 'OUTROS') if x.strip() else 'VOZ INCORRETA'
        )

        # Agrupar por 'ABA' e 'SITUAÇÃO MED.', contando as ocorrências
        grouped = df_filtered.groupby(['ABA', 'SITUAÇÃO MED.']).size().unstack(fill_value=0)

        # Preparar os dados para o gráfico
        aba_values = sorted(grouped.index)
        plot_data = grouped.T  # Transpor para ter categorias nas linhas e ABAs nas colunas

        # Plotar o gráfico de barras empilhadas
        width = 0.08  # largura das barras
        x = np.arange(len(aba_values))  # localização dos grupos no eixo x

        for i, (label, values) in enumerate(plot_data.iterrows()):
            ax.bar(x + i * width, values, width, label=label)

        # Configurações adicionais do gráfico
        ax.set_title('Contagem de Situação por ABA')
        ax.set_ylabel('Total')
        ax.set_xlabel('ABA')
        ax.set_xticks(x + width)
        ax.set_xticklabels(aba_values)

        # Adicionar rótulos com o total
        for i, (label, values) in enumerate(plot_data.iterrows()):
            for j, value in enumerate(values):
                if value > 0:
                    ax.text(x[j] + i * width, value, f'{int(value)}', ha='center', va='bottom')

        # Mostrar a legenda
        ax.legend()

        # Redesenhar o canvas do gráfico
        graph['canvas'].draw()

    def refresh_graph8(self, graph):
        ax = graph['ax']
        ax.clear()

        # Filtrar e agrupar os dados, ignorando valores vazios
        df_filtered = self.dataframe_cleaned.dropna(subset=['ABA', 'RESP MEDIÇÃO', 'SITUAÇÃO MED.'])

        # Verificar se o DataFrame filtrado está vazio
        if df_filtered.empty:
            ax.set_title('Nenhum dado disponível para exibir no gráfico')
            graph['canvas'].draw()
            return

        # Ordenar os valores de ABA
        aba_values = sorted(df_filtered['ABA'].unique())  # Ordenando os valores de ABA
        resp_medicao_values = df_filtered['RESP MEDIÇÃO'].unique()

        # Agrupar os dados por 'ABA', 'RESP MEDIÇÃO' e 'SITUAÇÃO MED.', e contar as ocorrências
        grouped = df_filtered.groupby(['ABA', 'RESP MEDIÇÃO'])['SITUAÇÃO MED.'].count().unstack(fill_value=0)

        # Verificar se há dados suficientes após o agrupamento
        if grouped.empty:
            ax.set_title('Nenhum dado disponível após o agrupamento')
            graph['canvas'].draw()
            return

        # Preparar o gráfico de barras empilhadas
        width = 0.35  # largura das barras
        x = np.arange(len(aba_values))  # localização dos grupos no eixo x

        # Inicializar o array de valores empilhados para cada RESP MEDIÇÃO
        bottom = np.zeros(len(aba_values))

        # Definir um conjunto de cores
        colors = plt.get_cmap('tab10').colors  # Usar um colormap padrão do matplotlib para cores

        # Iterar sobre cada 'RESP MEDIÇÃO'
        for i, resp in enumerate(resp_medicao_values):
            if resp in grouped.columns:
                # Extrair os valores da contagem por RESP MEDIÇÃO
                values = grouped[resp].reindex(aba_values, fill_value=0).values

                # Plotar as barras empilhadas
                bars = ax.bar(x, values, width, bottom=bottom, label=resp, color=colors[i % len(colors)])

                # Atualizar o 'bottom' para a próxima barra empilhada
                bottom += values

                # Adicionar rótulos de dados com a contagem
                for bar in bars:
                    height = bar.get_height()
                    if height > 0:
                        ax.text(bar.get_x() + bar.get_width() / 2, bar.get_y() + height / 2, f'{int(height)}', ha='center', va='center', fontsize=8, color='black')

        # Configurações adicionais do gráfico
        ax.set_title('Contagem de Situação por RESP MEDIÇÃO e ABA')
        ax.set_ylabel('Total de Situações')
        ax.set_xlabel('ABA')
        ax.set_xticks(x)
        ax.set_xticklabels(aba_values)

        # Mostrar a legenda
        ax.legend(title='RESP MEDIÇÃO')

        # Redesenhar o canvas do gráfico
        graph['canvas'].draw()

    def refresh_graph9(self, graph):
        """
        Atualiza o gráfico 9 com as configurações padrão definidas pelo usuário e adiciona uma linha em branco na tabela para separar da legenda.
        """

        # Limpar gráfico existente
        ax = graph['ax']
        ax.clear()

        # Corrigir o cálculo de valores para serem negativos
        aba_values = self.dataframe_cleaned['ABA'].unique().tolist()
        glosa_values = [-1 * value for value in self.dataframe_cleaned.groupby('ABA')['GLOSA - MANUTENÇÃO'].sum().tolist()]
        desc_com_values = [-1 * value for value in self.dataframe_cleaned.groupby('ABA')['DESC COMERCIAL'].sum().tolist()]
        km_exc_values = self.dataframe_cleaned.groupby('ABA')['KM EXCEDENTE'].sum().tolist()
        multa_values = self.dataframe_cleaned.groupby('ABA')['MULTA CONTRATUAL'].sum().tolist()
        ajustes_values = self.dataframe_cleaned.groupby('ABA')['AJUSTES / ACRÉCIMOS'].sum().tolist()

        x = np.arange(len(aba_values))  # Eixo X para os meses

        # Linhas para cada valor
        lines1, = ax.plot(x, glosa_values, marker='v', color='#ff7f0e', label='Glosa - Manutenção')
        lines2, = ax.plot(x, desc_com_values, marker='v', color='#d62728', label='Desc. Comercial')
        lines3, = ax.plot(x, km_exc_values, marker='^', color='#9467bd', label='KM Excedente')
        lines4, = ax.plot(x, multa_values, marker='^', color='#8c564b', label='Multa Contratual')
        lines5, = ax.plot(x, ajustes_values, marker='^', color='#e377c2', label='Ajustes / Acréscimos')

        # Adicionar a linha no valor zero
        ax.axhline(0, color='gray', linewidth=1, linestyle='--')  # Linha horizontal no valor zero

        # Ajustar a posição das linhas e dos rótulos no gráfico
        ax.set_xticks(x)
        ax.set_xticklabels(aba_values)

        # Adicionar legendas das linhas
        lines = [lines1, lines2, lines3, lines4, lines5]
        labels = [line.get_label() for line in lines]
        ax.legend(lines, labels, loc='upper right')

        # Formatação do eixo Y para valores monetários
        def currency_formatter(x, pos):
            return 'R$ {:,.2f}'.format(x).replace('.', 'x').replace(',', '.').replace('x', ',')

        ax.yaxis.set_major_formatter(FuncFormatter(currency_formatter))

        # Ajustar automaticamente os limites do eixo Y
        ax.set_ylim(auto=True)  # Remover limites manuais e ajustar automaticamente

        # Criar uma linha vazia (separadora)
        empty_line = ['' for _ in aba_values]

        # Criar a tabela de dados na parte inferior (sem cabeçalho)
        data = [empty_line,
                [f'R$ {value:,.2f}' for value in glosa_values],
                [f'R$ {value:,.2f}' for value in desc_com_values],
                [f'R$ {value:,.2f}' for value in km_exc_values],
                [f'R$ {value:,.2f}' for value in multa_values],
                [f'R$ {value:,.2f}' for value in ajustes_values]]

        # Títulos das linhas, incluindo a linha vazia
        row_labels = [''] + ['Glosa', 'Desc. Comercial', 'KM Excedente', 'Multa', 'Ajustes']

        # Adicionar a tabela ao gráfico (sem colLabels)
        table = ax.table(cellText=data, rowLabels=row_labels, cellLoc='center', loc='bottom')

        # Definir cores para cada linha
        colors = {
            'Glosa': '#ff7f0e',
            'Desc. Comercial': '#d62728',
            'KM Excedente': '#9467bd',
            'Multa': '#8c564b',
            'Ajustes': '#e377c2'
        }

        # Aplicar cores aos textos das células da tabela com base nos rótulos das linhas
        for i, row_label in enumerate(row_labels):
            if row_label in colors:
                for col in range(len(aba_values)):
                    cell = table[(i, col)]  # Acessar a célula usando (linha, coluna)
                    cell.set_text_props(color=colors[row_label])  # Definir a cor do texto

        # Definir valores automáticos de tamanho e espaçamento
        table_height = len(row_labels) * 0.05  # Calcula altura baseada no número de linhas
        table_scale_factor = 6  # Fator de escala para a altura da tabela
        table.scale(1, table_scale_factor)  # Aumenta a altura da tabela

        # Obter tamanho atual da figura
        fig = graph['canvas'].figure
        fig_width, fig_height = fig.get_size_inches()

        # Calcular altura total disponível para a tabela
        height_for_table = table_height * table_scale_factor
        bottom_margin = height_for_table / fig_height  # Porcentagem da altura para a margem inferior

        # Ajustar a altura do gráfico para o espaço disponível, incluindo a linha extra de separação
        fig.subplots_adjust(left=0.1, bottom=bottom_margin, right=0.95, top=0.85)  # Ajuste manual de subplots

        # Ajustar a altura das células da tabela
        for key, cell in table.get_celld().items():
            if key[1] == -1:  # Row Labels
                cell.set_text_props(weight='bold')
                cell.set_fontsize(9)
                cell.set_height(0.075)
            else:
                cell.set_fontsize(9)
                cell.set_height(0.075)  # Ajustar altura para caber corretamente

        # Desativar o cabeçalho (colLabels)
        table.auto_set_font_size(False)
        table.set_fontsize(9)

        # Redesenhar o gráfico e a tabela
        graph['canvas'].draw()

    def refresh_graph10(self, graph):
        """
        Novo gráfico baseado na coluna 'Nº CR' para visualizar o 'VALOR FATURADO' e 'PREV. MEDIÇÃO'.
        """
        ax = graph['ax']
        ax.clear()

        # Preparar os dados, garantindo que todos os 'Nº CR' sejam contabilizados, mesmo com dados ausentes
        cr_values = self.dataframe_cleaned['Nº CR'].unique()
        df_grouped = self.dataframe_cleaned.groupby('Nº CR').agg({
            'VALOR FATURADO': 'sum',
            'PREVISÃO DE MEDIÇÃO': 'sum',
            'GLOSA - MANUTENÇÃO': 'sum',
            'DESC COMERCIAL': 'sum',
            'KM EXCEDENTE': 'sum',
            'MULTA CONTRATUAL': 'sum',
            'AJUSTES / ACRÉCIMOS': 'sum'
        }).reindex(cr_values, fill_value=0)  # Preenche valores ausentes com 0

        # Calculando os valores de 'PREV. MEDIÇÃO'
        val_faturado = df_grouped['VALOR FATURADO']
        prev_medicao = (df_grouped['PREVISÃO DE MEDIÇÃO'] - df_grouped['GLOSA - MANUTENÇÃO'] - df_grouped['DESC COMERCIAL'] +
                        df_grouped['KM EXCEDENTE'] + df_grouped['MULTA CONTRATUAL'] + df_grouped['AJUSTES / ACRÉCIMOS'])

        x = np.arange(len(cr_values))  # Posição no eixo X para 'Nº CR'
        width = 0.35  # Largura das barras

        # Criar barras para o 'VALOR FATURADO' e 'PREV. MEDIÇÃO'
        bars1 = ax.bar(x - width / 2, val_faturado, width, label='Valor Faturado', color='#003f70')
        bars2 = ax.bar(x + width / 2, prev_medicao, width, label='Prev. Medição', color='#00afa0')

        # Rótulos para as barras, rotacionados em 90º
        for bar, value in zip(bars1, val_faturado):
            ax.text(bar.get_x() + bar.get_width() / 2, bar.get_height(), f'R$ {value:,.2f}', ha='center', va='bottom', rotation=90)
        for bar, value in zip(bars2, prev_medicao):
            ax.text(bar.get_x() + bar.get_width() / 2, bar.get_height(), f'R$ {value:,.2f}', ha='center', va='bottom', rotation=90)

        # Ajustes de labels e legendas
        ax.set_ylabel('Valores em R$')
        ax.set_xticks(x)
        ax.set_xticklabels(cr_values, rotation=45, ha='right')
        ax.legend()

        # Formatadores de moeda no eixo Y
        ax.yaxis.set_major_formatter(FuncFormatter(lambda x, _: f'R$ {x:,.2f}'))

        graph['canvas'].draw()

    def create_closures_page(self, page):
        closure_frame = ttk.Frame(page)
        closure_frame.pack(fill="both", expand=True, padx=20, pady=20)

        self.canvas = tk.Canvas(closure_frame)
        self.canvas.pack(side="left", fill="both", expand=True)

        self.scrollbar_y = ttk.Scrollbar(closure_frame, orient="vertical", command=self.canvas.yview)
        self.scrollbar_y.pack(side="right", fill="y")
        self.canvas.configure(yscrollcommand=self.scrollbar_y.set)

        self.card_frame = ttk.Frame(self.canvas)
        self.canvas.create_window((0, 0), window=self.card_frame, anchor="nw")

        self.refresh_closure_metrics()
        self.card_frame.bind("<Configure>", self.on_frame_configure)

    def on_frame_configure(self, event):
        self.canvas.configure(scrollregion=self.canvas.bbox("all"))

    def refresh_closure_metrics(self, event=None):
        for widget in self.card_frame.winfo_children():
            widget.destroy()

        df_filtered = self.dataframe_cleaned

        total_previsao_medicao = df_filtered['PREVISÃO DE MEDIÇÃO'].sum()
        total_valor_faturado = df_filtered['VALOR FATURADO'].sum()
        total_glosa = df_filtered['GLOSA - MANUTENÇÃO'].sum()
        total_desc_comercial = df_filtered['DESC COMERCIAL'].sum()
        total_multa_contratual = df_filtered['MULTA CONTRATUAL'].sum()
        total_km_excedente = df_filtered['KM EXCEDENTE'].sum()

        valid_rows = df_filtered.dropna(subset=['MEDIÇÃO EFETUADA', 'ENVIO FAT'])
        valid_rows = valid_rows[(pd.to_datetime(valid_rows['MEDIÇÃO EFETUADA'], dayfirst=True, errors='coerce').notna()) &
                                (pd.to_datetime(valid_rows['ENVIO FAT'], dayfirst=True, errors='coerce').notna())]
        valid_rows['DIFERENCA_DIAS'] = (pd.to_datetime(valid_rows['ENVIO FAT'], dayfirst=True) - pd.to_datetime(valid_rows['MEDIÇÃO EFETUADA'], dayfirst=True)).dt.days
        media_dias_geral = valid_rows['DIFERENCA_DIAS'].mean()

        total_medicoes_efetuadas = df_filtered['MEDIÇÃO EFETUADA'].dropna().count()
        total_medicoes_a_faturar = df_filtered['ENVIO FAT'].dropna().count()
        total_medicoes_finalizadas = df_filtered.dropna(subset=['ENVIO FAT', 'FAT MEDIÇÃO']).shape[0]

        self.create_card(
            self.card_frame, "Ano: 2024", 
            f"Previsão de Medição: R$ {total_previsao_medicao:,.2f}    "
            f"Valor Faturado: R$ {total_valor_faturado:,.2f}    "
            f"Média de Dias: {media_dias_geral:.1f} dias    \n"           
            f"Glosa: R$ {total_glosa:,.2f}    "
            f"Desconto: R$ {total_desc_comercial:,.2f}    "
            f"Multa: R$ {total_multa_contratual:,.2f}    "
            f"KM Excedente: R$ {total_km_excedente:,.2f}\n"
            f"Medições Efetuadas: {total_medicoes_efetuadas}/{df_filtered.shape[0]}    "
            f"Medições a Faturar: {total_medicoes_a_faturar}/{df_filtered.shape[0]}    "
            f"Medições Finalizadas: {total_medicoes_finalizadas}/{df_filtered.shape[0]}",
            "general"
        )

        for aba_value in df_filtered['ABA'].unique():
            aba_value_str = str(aba_value)
            if len(aba_value_str) < 4:
                continue

            aba_df = df_filtered[df_filtered['ABA'] == aba_value]
            mes = MESES.get(aba_value_str[-2:], 'Desconhecido')

            previsao_medicao = aba_df['PREVISÃO DE MEDIÇÃO'].sum()
            valor_faturado = aba_df['VALOR FATURADO'].sum()
            glosa = aba_df['GLOSA - MANUTENÇÃO'].sum()
            desc_comercial = aba_df['DESC COMERCIAL'].sum()
            multa_contratual = aba_df['MULTA CONTRATUAL'].sum()
            km_excedente = aba_df['KM EXCEDENTE'].sum()

            valid_rows_aba = aba_df.dropna(subset=['MEDIÇÃO EFETUADA', 'ENVIO FAT'])
            valid_rows_aba = valid_rows_aba[(pd.to_datetime(valid_rows_aba['MEDIÇÃO EFETUADA'], dayfirst=True, errors='coerce').notna()) &
                                            (pd.to_datetime(valid_rows_aba['ENVIO FAT'], dayfirst=True, errors='coerce').notna())]
            valid_rows_aba['DIFERENCA_DIAS'] = (pd.to_datetime(valid_rows_aba['ENVIO FAT'], dayfirst=True) - pd.to_datetime(valid_rows_aba['MEDIÇÃO EFETUADA'], dayfirst=True)).dt.days
            media_dias_aba = valid_rows_aba['DIFERENCA_DIAS'].mean()

            medicoes_efetuadas = aba_df['MEDIÇÃO EFETUADA'].dropna().count()
            medicoes_a_faturar = aba_df['ENVIO FAT'].dropna().count()
            medicoes_finalizadas = aba_df.dropna(subset=['ENVIO FAT', 'FAT MEDIÇÃO']).shape[0]

            self.create_card(
                self.card_frame, f"Mês {mes}:", 
                f"Previsão de Medição: R$ {previsao_medicao:,.2f}    "
                f"Valor Faturado: R$ {valor_faturado:,.2f}    "
                f"Média de Dias: {media_dias_aba:.1f} dias    \n"
                f"Glosa: R$ {glosa:,.2f}    "
                f"Desconto: R$ {desc_comercial:,.2f}    "
                f"Multa: R$ {multa_contratual:,.2f}    "
                f"KM Excedente: R$ {km_excedente:,.2f}\n"                
                f"Medições Efetuadas: {medicoes_efetuadas}/{aba_df.shape[0]}    "
                f"Medições a Faturar: {medicoes_a_faturar}/{aba_df.shape[0]}    "
                f"Medições Finalizadas: {medicoes_finalizadas}/{aba_df.shape[0]}",
                aba_value
            )

    def create_card(self, parent, title, content, tag):
        card = ttk.Frame(parent, relief="raise", borderwidth=2)
        card.pack(side="top", fill="x", padx=10, pady=5)

        title_frame = ttk.Frame(card)
        title_frame.pack(side="top", fill="x", padx=10, pady=5)

        title_label = ttk.Label(title_frame, text=title, font=("Roboto Mono", 14, "bold"), anchor="center")
        title_label.pack(side="left", expand=True)

        if tag != "general":
            button_frame = ttk.Frame(card)
            button_frame.pack(side="top", fill="x", padx=10, pady=5)
            
            expand_month_button = ttk.Button(button_frame, text="Expandir Mês", command=lambda: self.toggle_table(tag, "month"))
            expand_month_button.pack(side="right")

            expand_open_button = ttk.Button(button_frame, text="Expandir Abertos", command=lambda: self.toggle_table(tag, "open"))
            expand_open_button.pack(side="right")

            extract_report_button = ttk.Button(button_frame, text="Extrair Relatório(s)", command=lambda: self.extract_report(tag))
            self.table_frames[tag] = {
                "card": card,
                "month_frame": None,
                "open_frame": None,
                "extract_report_button": extract_report_button
            }

            extract_report_button.pack_forget()

        content_label = ttk.Label(card, text=self.format_content(content), font=("Roboto Mono", 10), anchor="center", justify="center")
        content_label.pack(side="top", padx=10, pady=5, fill="x")

        self.update_scrollregion()

    def format_content(self, content):
        lines = content.split('\n')
        max_length = max(len(line) for line in lines)
        formatted_lines = [line.ljust(max_length + 5) for line in lines]
        formatted_content = '\n'.join(formatted_lines)

        formatted_content = formatted_content.replace(',', 'TEMP')
        formatted_content = formatted_content.replace('.', ',')
        formatted_content = formatted_content.replace('TEMP', '.')

        return formatted_content

    def toggle_table(self, tag, table_type):
        table_frame_key = f"{table_type}_frame"
        if self.table_frames[tag][table_frame_key] is None:
            self.show_table(tag, table_type, table_frame_key)
        else:
            self.hide_table(tag, table_type, table_frame_key)
        self.update_scrollregion()
        self.update_extract_report_button(tag)

    def show_table(self, tag, table_type, table_frame_key):
        data = self.dataframe_cleaned[self.dataframe_cleaned['ABA'] == tag]
        if table_type == "open":
            data = data[data['ENVIO FAT'].isna() | data['FAT MEDIÇÃO'].isna()]

        columns = [
            "CLIENTE", "Nº MEDIÇÃO", "FECH. CONT.", "MEDIÇÃO EFETUADA",
            "APROV CLIENTE", "ENVIO FAT", "FAT MEDIÇÃO", "VALOR FATURADO", "PREVISÃO DE MEDIÇÃO",
            "RESP MEDIÇÃO", "SITUAÇÃO MED."
        ]

        column_widths = [500, 80, 80, 70, 70, 70, 70, 90, 90, 90, 90]

        table_frame = ttk.Frame(self.table_frames[tag]["card"])
        table_frame.pack(side="top", fill="x", padx=10, pady=10, expand=True)

        title = f"Mês {self.get_month_name(tag)} Expandido" if table_type == "month" else f"Mês {self.get_month_name(tag)} Abertos"
        title_label = ttk.Label(table_frame, text=title, font=("Roboto Mono", 12, "bold"), anchor="center")
        title_label.pack(side="top", pady=5, fill="x")

        tree_scroll_y = ttk.Scrollbar(table_frame, orient="vertical")
        tree_scroll_y.pack(side="right", fill="y")

        tree_scroll_x = ttk.Scrollbar(table_frame, orient="horizontal")
        tree_scroll_x.pack(side="bottom", fill="x")

        treeview = ttk.Treeview(table_frame, columns=columns, yscrollcommand=tree_scroll_y.set, xscrollcommand=tree_scroll_x.set, show="headings")
        treeview.pack(expand=True, fill="both")

        tree_scroll_y.config(command=treeview.yview)
        tree_scroll_x.config(command=treeview.xview)

        for col, width in zip(columns, column_widths):
            treeview.heading(col, text=col, anchor="w")
            treeview.column(col, width=width, anchor="w")

        for _, row in data.iterrows():
            values = [row[col] for col in columns]
            treeview.insert("", "end", values=values)

        self.table_frames[tag][table_frame_key] = table_frame

    def hide_table(self, tag, table_type, table_frame_key):
        if self.table_frames[tag][table_frame_key] is not None:
            self.table_frames[tag][table_frame_key].destroy()
            self.table_frames[tag][table_frame_key] = None

    def update_scrollregion(self):
        self.card_frame.update_idletasks()
        self.canvas.configure(scrollregion=self.canvas.bbox("all"))

    def update_extract_report_button(self, tag):
        if self.table_frames[tag]["month_frame"] or self.table_frames[tag]["open_frame"]:
            self.table_frames[tag]["extract_report_button"].pack(side="right")
        else:
            self.table_frames[tag]["extract_report_button"].pack_forget()

    def extract_report(self, tag):
        month_data = None
        open_data = None

        if self.table_frames[tag]["month_frame"]:
            month_data = self.dataframe_cleaned[self.dataframe_cleaned['ABA'] == tag]
        if self.table_frames[tag]["open_frame"]:
            open_data = self.dataframe_cleaned[(self.dataframe_cleaned['ABA'] == tag) & (self.dataframe_cleaned['ENVIO FAT'].isna() | self.dataframe_cleaned['FAT MEDIÇÃO'].isna())]

        if month_data is None and open_data is None:
            messagebox.showwarning("Aviso", "Nenhum relatório a ser extraído.")
            return

        now = datetime.now()
        current_time = now.strftime("%d-%m-%Y_%H-%M-%S")
        file_path = f"C:/Users/{os.getlogin()}/Desktop/Relatório Parcial ({self.get_month_name(tag)}) {current_time}.xlsx"

        with pd.ExcelWriter(file_path, engine='xlsxwriter') as writer:
            if month_data is not None:
                month_data.to_excel(writer, sheet_name='Relatório Mensal', index=False)
            if open_data is not None:
                open_data.to_excel(writer, sheet_name='Relatório Abertos', index=False)

        messagebox.showinfo("Extrair Relatório(s)", f"Relatório(s) extraído(s) com sucesso para {file_path}")

    def get_month_name(self, aba_value):
        aba_value_str = str(aba_value)
        return MESES.get(aba_value_str[-2:], 'Desconhecido')

    def refresh_graphs(self):
        for graph in self.graphs:
            # Executar o método de refresh para cada gráfico individualmente
            if graph['title'] == "Valores de Medição / Faturamento":
                self.refresh_graph1(graph)
            elif graph['title'] == "Clientes com Medição Aberta/Fechada por Mês":
                self.refresh_graph2(graph)
            elif graph['title'] == "Contagem de RESP MEDIÇÃO por Mês":
                self.refresh_graph3(graph)
            elif graph['title'] == "Novos Clientes vs Clientes Finalizados por Mês":
                self.refresh_graph4(graph)
            elif graph['title'] == "Total de Carros Locados por Mês":
                self.refresh_graph5(graph)
            elif graph['title'] == "Diferença Faturamento/Medição":
                self.refresh_graph6(graph)
            elif graph['title'] == "Contagem de Situação por ABA":
                self.refresh_graph7(graph)
            elif graph['title'] == "Contagem de Situação por PESSOA":
                self.refresh_graph8(graph)
            elif graph['title'] == "Valores Variáveis":
                self.refresh_graph9(graph)

            # Redesenhar o gráfico atualizado
            graph['canvas'].draw()

    def create_trainings_page(self, page):
        self.trainings_frame = ttk.Frame(page)
        self.trainings_frame.pack(fill="both", expand=True, padx=10, pady=10)

        training_files = self.get_training_files()
        for training_file in training_files:
            self.create_training_icon(training_file)

    def get_training_files(self):
        username = os.getlogin()
        file_paths = [
            f'C:\\Users\\{username}\\EBEC\\NC - Medicao - Documentos\\007 - TREINAMENTOS',
            f'C:\\Users\\joana.conceicao.EBEC-SA.000\\EBEC\\NC - Medicao - Documentos\\007 - TREINAMENTOS',
            f'C:\\Users\\{username}\\EBEC\\NC - Medicao - Documentos.000\\007 - TREINAMENTOS'
        ]

        training_files = []
        for file_path in file_paths:
            if os.path.exists(file_path):
                for file_name in os.listdir(file_path):
                    if file_name.lower().startswith("anexo") and file_name.lower().endswith(".pdf"):
                        training_files.append(os.path.join(file_path, file_name))
        return training_files

    def create_training_icon(self, training_file):
        file_name = os.path.basename(training_file)
        file_display_name = os.path.splitext(file_name)[0]

        icon_frame = ttk.Frame(self.trainings_frame, relief="raised", borderwidth=2)
        icon_frame.pack(side="top", fill="x", padx=10, pady=5)

        icon_label = ttk.Label(icon_frame, text=file_display_name, font=("Helvetica", 12), anchor="center")
        icon_label.pack(side="left", padx=10, pady=5)

        open_button = ttk.Button(icon_frame, text="Abrir", command=lambda: self.open_file(training_file))
        open_button.pack(side="right", padx=10, pady=5)

    def open_file(self, file_path):
        webbrowser.open(f'file://{file_path}')

    def create_comparison_page(self, page):
        self.create_title(page, "Comparar Meses")

        comparison_frame = ttk.Frame(page)
        comparison_frame.pack(fill="x", padx=10, pady=10)

        months = self.dataframe_cleaned['ABA'].unique().tolist()

        self.month1_var = tk.StringVar()
        self.month2_var = tk.StringVar()

        ttk.Label(comparison_frame, text="Mês 1:").grid(row=0, column=0, padx=5, pady=5, sticky="e")
        self.month1_select = AutocompleteCombobox(comparison_frame, textvariable=self.month1_var, width=55)
        self.month1_select.set_completion_list(months)
        self.month1_select.grid(row=0, column=1, padx=5, pady=5, sticky="w")

        ttk.Label(comparison_frame, text="Mês 2:").grid(row=0, column=2, padx=5, pady=5, sticky="e")
        self.month2_select = AutocompleteCombobox(comparison_frame, textvariable=self.month2_var, width=55)
        self.month2_select.set_completion_list(months)
        self.month2_select.grid(row=0, column=3, padx=5, pady=5, sticky="w")

        compare_button = ttk.Button(comparison_frame, text="Comparar", command=self.compare_months)
        compare_button.grid(row=0, column=4, columnspan=2, padx=5, pady=5)

        self.comparison_treeview = ttk.Treeview(page, show="headings")
        self.comparison_treeview.pack(fill="both", expand=True, padx=10, pady=10)

        export_button = ttk.Button(page, text="Exportar Comparação", command=self.export_comparison)
        export_button.pack(fill="x", padx=10, pady=10)

    def compare_months(self):
        epsilon = 0.05  # Margem de erro para desconsiderar variações de até 1 centavo
        month1 = self.month1_var.get().strip()
        month2 = self.month2_var.get().strip()

        if not month1 or not month2:
            messagebox.showwarning("Aviso", "Por favor, preencha os dois meses para a comparação.")
            return

        df_mes1 = self.dataframe_cleaned[self.dataframe_cleaned['ABA'] == month1].copy()
        df_mes2 = self.dataframe_cleaned[self.dataframe_cleaned['ABA'] == month2].copy()

        if df_mes1.empty:
            messagebox.showerror("Erro", f"Não foram encontrados dados para o mês {month1}.")
            return
        if df_mes2.empty:
            messagebox.showerror("Erro", f"Não foram encontrados dados para o mês {month2}.")
            return

        df_mes1['Código'] = df_mes1['Nº MEDIÇÃO'].apply(lambda x: x.split('-')[0])
        df_mes2['Código'] = df_mes2['Nº MEDIÇÃO'].apply(lambda x: x.split('-')[0])

        colunas_comparar = [
            'PREVISÃO DE MEDIÇÃO', 'GLOSA - MANUTENÇÃO', 'DESC COMERCIAL', 
            'KM EXCEDENTE', 'MULTA CONTRATUAL', 'AJUSTES / ACRÉCIMOS',
            'QTDE LOCADOS', 'QTDE RESERVA'
        ]

        resultados = []

        codigos = set(df_mes1['Código'].unique()).union(set(df_mes2['Código'].unique()))

        for codigo in codigos:
            linha_mes1 = df_mes1[df_mes1['Código'] == codigo].iloc[0] if codigo in df_mes1['Código'].values else None
            linha_mes2 = df_mes2[df_mes2['Código'] == codigo].iloc[0] if codigo in df_mes2['Código'].values else None

            if linha_mes1 is None:
                diff = {'ABA': month2, 'CLIENTE': linha_mes2['CLIENTE'], 'Nº MEDIÇÃO': codigo}
                for coluna in colunas_comparar:
                    diff[coluna] = -pd.Series(linha_mes2[coluna]).fillna(0).values[0]
                diff['Valor Diferença'] = (
                    -pd.Series(linha_mes2['PREVISÃO DE MEDIÇÃO']).fillna(0).values[0]
                    - pd.Series(linha_mes2['GLOSA - MANUTENÇÃO']).fillna(0).values[0]
                    - pd.Series(linha_mes2['DESC COMERCIAL']).fillna(0).values[0]
                    + pd.Series(linha_mes2['KM EXCEDENTE']).fillna(0).values[0]
                    + pd.Series(linha_mes2['MULTA CONTRATUAL']).fillna(0).values[0]
                    + pd.Series(linha_mes2['AJUSTES / ACRÉCIMOS']).fillna(0).values[0]
                )
            elif linha_mes2 is None:
                diff = {'ABA': month1, 'CLIENTE': linha_mes1['CLIENTE'], 'Nº MEDIÇÃO': codigo}
                for coluna in colunas_comparar:
                    diff[coluna] = pd.Series(linha_mes1[coluna]).fillna(0).values[0]
                diff['Valor Diferença'] = (
                    pd.Series(linha_mes1['PREVISÃO DE MEDIÇÃO']).fillna(0).values[0]
                    - pd.Series(linha_mes1['GLOSA - MANUTENÇÃO']).fillna(0).values[0]
                    - pd.Series(linha_mes1['DESC COMERCIAL']).fillna(0).values[0]
                    + pd.Series(linha_mes1['KM EXCEDENTE']).fillna(0).values[0]
                    + pd.Series(linha_mes1['MULTA CONTRATUAL']).fillna(0).values[0]
                    + pd.Series(linha_mes1['AJUSTES / ACRÉCIMOS']).fillna(0).values[0]
                )
            else:
                diff = {
                    'ABA': f"{month1} vs {month2}", 'CLIENTE': linha_mes1['CLIENTE'], 'Nº MEDIÇÃO': codigo
                }
                for coluna in colunas_comparar:
                    diff[coluna] = (
                        pd.Series(linha_mes2[coluna]).fillna(0).values[0] - 
                        pd.Series(linha_mes1[coluna]).fillna(0).values[0]
                    )
                diff['Valor Diferença'] = (
                    (pd.Series(linha_mes2['PREVISÃO DE MEDIÇÃO']).fillna(0).values[0] - pd.Series(linha_mes1['PREVISÃO DE MEDIÇÃO']).fillna(0).values[0]) 
                    - (pd.Series(linha_mes2['GLOSA - MANUTENÇÃO']).fillna(0).values[0] - pd.Series(linha_mes1['GLOSA - MANUTENÇÃO']).fillna(0).values[0])
                    - (pd.Series(linha_mes2['DESC COMERCIAL']).fillna(0).values[0] - pd.Series(linha_mes1['DESC COMERCIAL']).fillna(0).values[0])
                    + (pd.Series(linha_mes2['KM EXCEDENTE']).fillna(0).values[0] - pd.Series(linha_mes1['KM EXCEDENTE']).fillna(0).values[0])
                    + (pd.Series(linha_mes2['MULTA CONTRATUAL']).fillna(0).values[0] - pd.Series(linha_mes1['MULTA CONTRATUAL']).fillna(0).values[0])
                    + (pd.Series(linha_mes2['AJUSTES / ACRÉCIMOS']).fillna(0).values[0] - pd.Series(linha_mes1['AJUSTES / ACRÉCIMOS']).fillna(0).values[0])
                )
                
            # Aplicar margem de erro
            diff['Valor Diferença'] = 0 if abs(diff['Valor Diferença']) <= epsilon else diff['Valor Diferença']

            resultados.append(diff)

        df_resultados = pd.DataFrame(resultados)

        df_resultados = df_resultados.round(2)
        df_resultados = df_resultados.loc[:, (df_resultados != 0).any(axis=0)]

        colunas_final = ['ABA', 'CLIENTE', 'Nº MEDIÇÃO', 'Valor Diferença'] + colunas_comparar
        colunas_disponiveis = [col for col in colunas_final if col in df_resultados.columns]
        df_resultados = df_resultados[colunas_disponiveis]

        self.df_resultados = df_resultados

        self.populate_comparison_treeview(df_resultados)

    def populate_comparison_treeview(self, dataframe):
        self.comparison_treeview.delete(*self.comparison_treeview.get_children())
        self.comparison_treeview["columns"] = dataframe.columns.tolist()
        for col in self.comparison_treeview["columns"]:
            self.comparison_treeview.heading(col, text=col, anchor=tk.W)
            self.comparison_treeview.column(col, anchor=tk.W)
        for index, row in dataframe.iterrows():
            self.comparison_treeview.insert("", "end", values=row.tolist())

    def export_comparison(self):
        if not hasattr(self, 'df_resultados'):
            messagebox.showwarning("Aviso", "Nenhum dado de comparação disponível para exportar.")
            return

        now = datetime.now()
        current_time = now.strftime("%d-%m-%Y_%H-%M-%S")
        file_path = f"C:/Users/{os.getlogin()}/Desktop/Comparação_Meses_{current_time}.xlsx"

        self.df_resultados.to_excel(file_path, index=False)
        messagebox.showinfo("Exportar Comparação", f"Comparação exportada com sucesso para {file_path}")

    def create_verification_page(self, page):
        """Cria a aba de Verificação de Faturamento com funcionalidades de filtro, ordenação e exportação."""
        self.create_title(page, "Verificação de Faturamento")
        self.create_verification_buttons(page)
        self.create_verification_treeview(page)

    def sort_status_column(self, col):
        """Ordena a Treeview com base na coluna especificada."""
        try:
            data = [(self.status_treeview.set(child, col), child) for child in self.status_treeview.get_children("")]
            data.sort(reverse=False)
            for i, (_, child) in enumerate(data):
                self.status_treeview.move(child, "", i)
        except Exception as e:
            print(f"Erro ao ordenar a coluna {col}: {e}")

    def create_verification_buttons(self, page):
        button_frame = ttk.Frame(page)
        button_frame.pack(side="top", fill="x", padx=10, pady=5)

        generate_report_button = ttk.Button(button_frame, text="Gerar Relatório de Verificação", command=self.generate_verification_report)
        generate_report_button.pack(side="left", padx=5, pady=5)

    def create_verification_treeview(self, page):
        tree_frame = ttk.Frame(page)
        tree_frame.pack(fill="both", expand=True)

        tree_scroll_y = ttk.Scrollbar(tree_frame, orient="vertical")
        tree_scroll_y.pack(side="right", fill="y")

        tree_scroll_x = ttk.Scrollbar(tree_frame, orient="horizontal")
        tree_scroll_x.pack(side="bottom", fill="x")

        self.verification_treeview = ttk.Treeview(tree_frame, yscrollcommand=tree_scroll_y.set, xscrollcommand=tree_scroll_x.set, show="headings")
        self.verification_treeview.pack(expand=True, fill="both")

        tree_scroll_y.config(command=self.verification_treeview.yview)
        tree_scroll_x.config(command=self.verification_treeview.xview)

        self.populate_verification_treeview(self.verification_dataframe)

    def setup_verification_treeview_columns(self, dataframe):
        """Configura as colunas do Treeview e adiciona eventos de clique para ordenar as colunas."""
        self.verification_treeview["columns"] = dataframe.columns.tolist()
        default_font = Font()
        for col in self.verification_treeview["columns"]:
            self.verification_treeview.heading(col, text=col, anchor=tk.W, command=lambda c=col: self.sort_verification_column(c))
            column_width = default_font.measure(col.title())
            self.verification_treeview.column(col, width=column_width, stretch=False)

    def populate_verification_treeview(self, dataframe):
        """Preenche a Treeview de verificação de faturamento com os dados do dataframe."""
        self.verification_treeview.delete(*self.verification_treeview.get_children())
        self.setup_verification_treeview_columns(dataframe)
        for index, row in dataframe.iterrows():
            self.verification_treeview.insert("", "end", values=row.tolist(), tags=(index,))
        self.adjust_verification_column_widths()

    def adjust_verification_column_widths(self):
        """Ajusta as larguras das colunas do Treeview de verificação para se adequar ao conteúdo."""
        for col in self.verification_treeview["columns"]:
            max_width = Font().measure(col.title())
            for row in self.verification_treeview.get_children():
                row_value = self.verification_treeview.item(row, 'values')[self.verification_treeview["columns"].index(col)]
                row_width = Font().measure(str(row_value))
                if row_width > max_width:
                    max_width = row_width
            self.verification_treeview.column(col, width=max_width)

    def sort_verification_column(self, col):
        """Ordena a coluna especificada na Treeview de verificação."""
        try:
            data = [(self.verification_treeview.set(child, col), child) for child in self.verification_treeview.get_children("")]
            # Verifica se o texto no cabeçalho possui uma seta indicando a ordenação
            header_text = self.verification_treeview.heading(col)["text"]
            reverse = header_text[-1] == '↓'  # Inverte a ordenação se estiver ordenado de forma crescente

            # Realiza a ordenação com base no valor do item
            data.sort(reverse=reverse)

            for index, (_, child) in enumerate(data):
                self.verification_treeview.move(child, "", index)

            # Atualiza o texto do cabeçalho para indicar a direção da ordenação
            new_heading_text = f"{header_text[:-1]}↓" if not reverse else f"{header_text[:-1]}↑"
            self.verification_treeview.heading(col, text=new_heading_text)
        except Exception as e:
            messagebox.showerror("Erro", f"Erro ao ordenar a coluna '{col}': {e}")

    def generate_verification_report(self):
        if self.verification_dataframe.empty:
            messagebox.showwarning("Aviso", "Nenhuma diferença encontrada para gerar o relatório.")
            return

        now = datetime.now()
        current_time = now.strftime("%d-%m-%Y_%H-%M-%S")
        file_path = f"C:/Users/{os.getlogin()}/Desktop/Relatório_de_Verificação_{current_time}.xlsx"

        self.verification_dataframe.to_excel(file_path, index=False)
        messagebox.showinfo("Relatório de Verificação", f"Relatório de verificação gerado com sucesso para {file_path}")

    def create_status_tracking_page(self, page):
        """Cria a aba de acompanhamento de status de clientes por mês."""
        self.create_title(page, "Acompanhamento de Status de Clientes por Mês")

        # Criação da matriz (nmed|cliente | 2401 | 2402 | ... | 2412)
        matrix_frame = ttk.Frame(page)
        matrix_frame.pack(fill="both", expand=True, padx=10, pady=10)

        # Gerar a matriz de status
        status_matrix = self.generate_status_matrix()
        columns = ['Nº Medição', 'Cliente'] + [f'24{str(i).zfill(2)}' for i in range(1, 13)]

        # Exibir a matriz de status em uma Treeview
        tree_scroll_y = ttk.Scrollbar(matrix_frame, orient="vertical")
        tree_scroll_y.pack(side="right", fill="y")

        tree_scroll_x = ttk.Scrollbar(matrix_frame, orient="horizontal")
        tree_scroll_x.pack(side="bottom", fill="x")

        self.status_treeview = ttk.Treeview(matrix_frame, columns=columns, yscrollcommand=tree_scroll_y.set, xscrollcommand=tree_scroll_x.set, show="headings")
        self.status_treeview.pack(expand=True, fill="both")

        tree_scroll_y.config(command=self.status_treeview.yview)
        tree_scroll_x.config(command=self.status_treeview.xview)

        # Ajustar as larguras das colunas
        self.status_treeview.column('Nº Medição', width=100, anchor=tk.W)  # Coluna dos 4 dígitos
        self.status_treeview.column('Cliente', width=300, anchor=tk.W)  # Coluna do cliente
        for col in columns[2:]:
            self.status_treeview.column(col, width=50, anchor=tk.W)  # Cada mês com largura 50

        for col in columns:
            self.status_treeview.heading(col, text=col, anchor=tk.W)

        # Definir a tag para o status "P" com a cor desejada
        self.status_treeview.tag_configure('pendente', background='#FFFF99')  # Definir a cor de fundo amarela

        for row in status_matrix:
            # Inserir a linha na Treeview e obter o row_id para referência
            row_id = self.status_treeview.insert("", "end", values=row)
            
            # Iterar por cada valor na linha para identificar o status "P"
            for col_index, value in enumerate(row):
                if value == 'P':
                    # Aplique a tag 'pendente' ao valor 'P'
                    self.status_treeview.tag_configure('pendente', background='#FFFF99')  # Cor amarela para "P"
                    # Tag para a célula específica na coluna que contém "P"
                    self.status_treeview.item(row_id, tags=('pendente',))

        # Definir uma cor diferente para as linhas alternadas
        self.status_treeview.tag_configure('evenrow', background='#f2f2f2')
        self.status_treeview.tag_configure('oddrow', background='white')
        for index, item in enumerate(self.status_treeview.get_children()):
            if index % 2 == 0:
                self.status_treeview.item(item, tags=('evenrow',))
            else:
                self.status_treeview.item(item, tags=('oddrow',))

    def create_status_tracking_page(self, page):
        """Cria a aba de acompanhamento de status de clientes por mês."""
        self.create_title(page, "Acompanhamento de Status de Clientes por Mês")

        # Criação do frame de filtros
        filter_frame = ttk.Frame(page)
        filter_frame.pack(side="top", fill="x", padx=10, pady=10)

        # Filtros de status (checkbox)
        self.status_filter_a = tk.BooleanVar(value=True)
        self.status_filter_p = tk.BooleanVar(value=True)
        self.status_filter_x = tk.BooleanVar(value=True)

        self.filter_checkbutton_a = ttk.Checkbutton(filter_frame, text="A (Ativo)", variable=self.status_filter_a)
        self.filter_checkbutton_a.pack(side="left", padx=5)

        self.filter_checkbutton_p = ttk.Checkbutton(filter_frame, text="P (Pendente)", variable=self.status_filter_p)
        self.filter_checkbutton_p.pack(side="left", padx=5)

        self.filter_checkbutton_x = ttk.Checkbutton(filter_frame, text="X (Encerrado)", variable=self.status_filter_x)
        self.filter_checkbutton_x.pack(side="left", padx=5)

        self.apply_filter_button = ttk.Button(filter_frame, text="FILTRAR", command=self.apply_status_filter)
        self.apply_filter_button.pack(side="left", padx=10)

        # Botão de extração da tabela
        self.export_button = ttk.Button(filter_frame, text="Exportar Tabela", command=self.export_status_table)
        self.export_button.pack(side="right", padx=10)

        # Criação da matriz de status
        matrix_frame = ttk.Frame(page)
        matrix_frame.pack(fill="both", expand=True, padx=10, pady=10)

        # Gerar a matriz de status com a nova coluna "RESP MEDIÇÃO"
        status_matrix = self.generate_status_matrix_with_resp_medicao()
        columns = ['Nº Medição', 'Cliente', 'RESP MEDIÇÃO'] + [f'24{str(i).zfill(2)}' for i in range(1, 13)]

        # Exibir a matriz de status em uma Treeview
        tree_scroll_y = ttk.Scrollbar(matrix_frame, orient="vertical")
        tree_scroll_y.pack(side="right", fill="y")

        tree_scroll_x = ttk.Scrollbar(matrix_frame, orient="horizontal")
        tree_scroll_x.pack(side="bottom", fill="x")

        self.status_treeview = ttk.Treeview(matrix_frame, columns=columns, yscrollcommand=tree_scroll_y.set, xscrollcommand=tree_scroll_x.set, show="headings")
        self.status_treeview.pack(expand=True, fill="both")

        tree_scroll_y.config(command=self.status_treeview.yview)
        tree_scroll_x.config(command=self.status_treeview.xview)

        # Ajustar as larguras das colunas
        self.status_treeview.column('Nº Medição', width=100, anchor=tk.W)  # Coluna dos 4 dígitos
        self.status_treeview.column('Cliente', width=300, anchor=tk.W)  # Coluna do cliente
        self.status_treeview.column('RESP MEDIÇÃO', width=150, anchor=tk.W)  # Coluna "RESP MEDIÇÃO"
        for col in columns[3:]:
            self.status_treeview.column(col, width=50, anchor=tk.W)  # Cada mês com largura 50

        for col in columns:
            # Adiciona o evento de clique no cabeçalho para ordenar a coluna correspondente
            self.status_treeview.heading(col, text=col, anchor=tk.W, command=lambda c=col: self.sort_status_column(c))

        # Inserir os valores e aplicar as tags de coloração
        for row in status_matrix:
            # Inserir a linha na Treeview
            row_id = self.status_treeview.insert("", "end", values=row)

            # Aplicar as tags de coloração baseadas nos status
            if 'P' in row:
                self.status_treeview.item(row_id, tags=('pendente',))
            if 'X' in row:
                self.status_treeview.item(row_id, tags=('encerrado',))

        # Definir as tags para as cores
        self.status_treeview.tag_configure('pendente', background='#FFCCCC')  # Vermelho claro para linhas com "P".
        self.status_treeview.tag_configure('encerrado', background='#D3D3D3')  # Cinza para linhas com "X".

        # Definir uma cor diferente para as linhas alternadas
        self.status_treeview.tag_configure('evenrow', background='#f2f2f2')
        self.status_treeview.tag_configure('oddrow', background='white')
        for index, item in enumerate(self.status_treeview.get_children()):
            if index % 2 == 0:
                self.status_treeview.item(item, tags=('evenrow',))
            else:
                self.status_treeview.item(item, tags=('oddrow',))

    def generate_status_matrix_with_resp_medicao(self):
        """Gera a matriz de status com a coluna 'RESP MEDIÇÃO' baseada no dataframe de controle."""
        # Selecionar colunas relevantes
        df_status = self.dataframe_cleaned[['CLIENTE', 'ABA', 'STATUS', 'Nº MEDIÇÃO', 'RESP MEDIÇÃO']].copy()
        df_status = df_status.dropna(subset=['CLIENTE', 'ABA', 'STATUS', 'Nº MEDIÇÃO'])

        # Gerar lista de clientes únicos
        clientes = df_status['CLIENTE'].unique().tolist()
        meses = [f'24{str(i).zfill(2)}' for i in range(1, 13)]

        # Inicializar a matriz
        status_matrix = []

        for cliente in clientes:
            cliente_row = []
            cliente_data = df_status[df_status['CLIENTE'] == cliente]

            # Obter os 4 primeiros dígitos da primeira ocorrência de 'Nº MEDIÇÃO'
            n_medicao = cliente_data['Nº MEDIÇÃO'].iloc[0].split('-')[0][:4]

            # Adicionar o código de medição e o nome do cliente
            cliente_row.append(n_medicao)  # Nova coluna com os 4 primeiros dígitos
            cliente_row.append(cliente)

            # Obter o responsável pela medição mais recente para o mês atual
            resp_medicao = ""
            for mes in meses[::-1]:
                mes_data = cliente_data[cliente_data['ABA'] == mes]
                if not mes_data.empty:
                    resp_medicao = mes_data['RESP MEDIÇÃO'].iloc[0]
                    break

            # Adicionar a coluna "RESP MEDIÇÃO"
            cliente_row.append(resp_medicao)

            cliente_status = {str(aba): status for aba, status in zip(cliente_data['ABA'], cliente_data['STATUS'])}

            ultimo_status = None
            primeiro_pendente = False  # Indica se o primeiro P já foi encontrado

            for mes in meses:
                if mes in cliente_status:
                    status_atual = cliente_status[mes].strip().upper()

                    if status_atual in ['ATIVO', 'AG. FAT.', 'PARCIAL']:
                        cliente_row.append('A')  # Aberto (Ativo)
                        ultimo_status = 'A'
                    elif status_atual == 'FINALIZADO':
                        cliente_row.append('F')  # Fechado (Finalizado)
                        ultimo_status = 'F'
                    else:
                        cliente_row.append('O')  # Outros status, mapeados como "O"
                        ultimo_status = 'O'
                else:
                    if ultimo_status == 'A' and not primeiro_pendente:
                        cliente_row.append('P')  # Pendente
                        primeiro_pendente = True
                        ultimo_status = 'P'
                    elif ultimo_status == 'F':
                        cliente_row.append('X')  # Inativo após Fechado (Finalizado)
                    elif primeiro_pendente:
                        cliente_row.append('')  # Vazio após o primeiro P
                    else:
                        cliente_row.append('')  # Nenhum status conhecido antes de P

            status_matrix.append(cliente_row)

        return status_matrix

    def export_status_table(self):
        """Exporta a tabela de acompanhamento de status para um arquivo Excel."""
        # Criar um dataframe com as colunas e linhas da Treeview
        columns = [self.status_treeview.heading(col)["text"] for col in self.status_treeview["columns"]]
        data = [self.status_treeview.item(item)["values"] for item in self.status_treeview.get_children()]

        df_export = pd.DataFrame(data, columns=columns)

        # Gerar o nome do arquivo com data e hora
        now = datetime.now().strftime("%Y-%m-%d_%H-%M-%S")
        file_path = f"C:/Users/{os.getlogin()}/Desktop/Acomp_Status_Abertos_{now}.xlsx"

        # Salvar o dataframe em um arquivo Excel
        df_export.to_excel(file_path, index=False)
        messagebox.showinfo("Exportar Tabela", f"Tabela exportada com sucesso para {file_path}")


    def populate_status_treeview(self, matrix):
        """Popula a Treeview com os dados da matriz filtrada."""
        # Limpar a Treeview antes de popular novamente
        self.status_treeview.delete(*self.status_treeview.get_children())

        # Definir as colunas da Treeview (mantém as existentes)
        columns = ['Nº Medição', 'Cliente', 'RESP MEDIÇÃO'] + [f'24{str(i).zfill(2)}' for i in range(1, 13)]
        self.status_treeview["columns"] = columns

        # Redefinir os títulos das colunas na Treeview
        for col in columns:
            self.status_treeview.heading(col, text=col, anchor=tk.W)
            self.status_treeview.column(col, width=100 if col in ['Nº Medição', 'Cliente', 'RESP MEDIÇÃO'] else 50, anchor=tk.W)

        # Inserir os dados da matriz na Treeview
        for row in matrix:
            # Adiciona a linha completa
            row_id = self.status_treeview.insert("", "end", values=row)

            # Adiciona tags específicas para linhas contendo status 'P' ou 'X'
            if 'P' in row:
                self.status_treeview.item(row_id, tags=('pendente',))
            if 'X' in row:
                self.status_treeview.item(row_id, tags=('encerrado',))

        # Reaplicar coloração das linhas
        self.reapply_row_coloring()

    def reassign_sorting_to_columns(self):
        """
        Reaplica o comando de ordenação para cada coluna da Treeview após a aplicação do filtro.
        """
        for col in self.status_treeview["columns"]:
            self.status_treeview.heading(col, command=lambda c=col: self.sort_status_column(c))

    def apply_status_filter(self):
        """Aplica o filtro de status baseado no mês atual e nos filtros selecionados."""
        current_month = f"24{datetime.now().month:02d}"

        filter_a = self.status_filter_a.get()
        filter_p = self.status_filter_p.get()
        filter_x = self.status_filter_x.get()

        for item in self.status_treeview.get_children():
            self.status_treeview.delete(item)

        status_matrix = self.generate_status_matrix_with_resp_medicao()
        filtered_matrix = []
        for row in status_matrix:
            status_values = row[3:]
            month_index = int(current_month[-2:]) - 1

            status_atual = status_values[month_index]
            if (status_atual == 'A' and filter_a) or (status_atual == 'P' and filter_p) or (status_atual == 'X' and filter_x):
                filtered_matrix.append(row)

        self.populate_status_treeview(filtered_matrix)

        # Reaplicar a coloração das linhas
        self.reapply_row_coloring()

        # Reaplicar a configuração de ordenação após aplicar o filtro
        self.reassign_sorting_to_columns()

    def reapply_row_coloring(self):
        """Reaplica a coloração das linhas e a alternância."""
        for index, item in enumerate(self.status_treeview.get_children()):
            # Alternar entre linhas pares e ímpares para aplicar cor
            if index % 2 == 0:
                self.status_treeview.item(item, tags=('evenrow',))
            else:
                self.status_treeview.item(item, tags=('oddrow',))

    def apply_color_to_cell(self, row_id, col_index, color):
        """Aplica a cor à célula específica com base no index da coluna."""
        # Obter o valor atual da célula para mantê-lo intacto
        current_value = self.status_treeview.item(row_id, "values")[col_index]

        # Redefinir o valor com a cor aplicada
        self.status_treeview.item(row_id, values=(
            *self.status_treeview.item(row_id, "values")[:col_index],
            f'{{{current_value}}}',  # Adicionar chaves ao redor para destacar que é um valor formatado (opcional)
            *self.status_treeview.item(row_id, "values")[col_index + 1:]
        ))

        # Aplicar a tag 'pendente' à célula com o status 'P'
        self.status_treeview.tag_configure('pendente', background=color)  # Define a cor amarela para "P"
        self.status_treeview.item(row_id, tags=('pendente',))

if __name__ == "__main__":
    username = os.getlogin()
    file_paths = [
        f'C:\\Users\\{username}\\EBEC\\NC - Medicao - Documentos\\003 - POWER BI MEDIÇÃO\\001 - RELATORIOS BI\\001 - Fechamento\\RELATORIO GERAL MEDIÇÃO.xlsx',
        f'C:\\Users\\joana.conceicao.EBEC-SA.000\\EBEC\\NC - Medicao - Documentos\\003 - POWER BI MEDIÇÃO\\001 - RELATORIOS BI\\001 - Fechamento\\RELATORIO GERAL MEDIÇÃO.xlsx',
        f'C:\\Users\\{username}\\EBEC\\NC - Medicao - Documentos.000\\003 - POWER BI MEDIÇÃO\\001 - RELATORIOS BI\\001 - Fechamento\\RELATORIO GERAL MEDIÇÃO.xlsx'
    ]

    for file_path in file_paths:
        if os.path.exists(file_path):
            df = pd.read_excel(file_path)
            break
    else:
        file_path = filedialog.askopenfilename(title="Selecione o arquivo RELATORIO GERAL MEDIÇÃO", filetypes=[("Excel files", "*.xlsx")])
        if not file_path:
            messagebox.showwarning("Aviso", "Arquivo não selecionado.")
            exit()
        df = pd.read_excel(file_path)

    viewer = DataFrameViewer(df)
    viewer.mainloop()
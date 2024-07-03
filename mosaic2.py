import tkinter as tk
from tkinter import simpledialog, messagebox, ttk
import pandas as pd
import matplotlib.pyplot as plt
import seaborn as sns
from matplotlib.backends.backend_tkagg import FigureCanvasTkAgg
from datetime import datetime
import os


class DashboardApp:
    def __init__(self, root):
        self.root = root
        self.root.title("Granulometria Mosaic")
        self.root.geometry("800x600")

        self.create_widgets()

    def create_widgets(self):
        # Título
        title_label = tk.Label(self.root, text="Granulometria Mosaic", font=("Helvetica", 24, "bold"), pady=10)
        title_label.pack()

        # Formulário de entrada
        input_frame = ttk.LabelFrame(self.root, text="Entrada de Dados")
        input_frame.pack(pady=20, padx=10, fill="both")

        nome_label = tk.Label(input_frame, text="Nome:")
        nome_label.grid(row=0, column=0, padx=10, pady=5, sticky="w")
        self.nome_entry = ttk.Entry(input_frame)
        self.nome_entry.grid(row=0, column=1, padx=10, pady=5)
        self.nome_entry.bind("<Return>", lambda event: self.focus_next(event, self.data_entry))

        data_label = tk.Label(input_frame, text="Data (DD/MM/AAAA):")
        data_label.grid(row=1, column=0, padx=10, pady=5, sticky="w")
        self.data_entry = ttk.Entry(input_frame)
        self.data_entry.grid(row=1, column=1, padx=10, pady=5)
        self.data_entry.bind("<KeyRelease>", self.autocomplete_date)  # Autocomplete date with '/'
        self.data_entry.bind("<Return>", lambda event: self.focus_next(event, self.hora_entry))  # Avança para a hora

        hora_label = tk.Label(input_frame, text="Hora (HH:MM):")
        hora_label.grid(row=2, column=0, padx=10, pady=5, sticky="w")
        self.hora_entry = ttk.Entry(input_frame)
        self.hora_entry.grid(row=2, column=1, padx=10, pady=5)
        self.hora_entry.bind("<KeyRelease>", self.autocomplete_time)  # Autocomplete time with ':'
        self.hora_entry.bind("<Return>",
                             lambda event: self.focus_next(event, self.malha_7_entry))  # Avança para a próxima entrada

        malha_7_label = tk.Label(input_frame, text="Valor da Malha 7:")
        malha_7_label.grid(row=3, column=0, padx=10, pady=5, sticky="w")
        self.malha_7_entry = ttk.Entry(input_frame)
        self.malha_7_entry.grid(row=3, column=1, padx=10, pady=5)
        self.malha_7_entry.bind("<Return>", lambda event: self.focus_next(event, self.malha_8_entry))

        malha_8_label = tk.Label(input_frame, text="Valor da Malha 8:")
        malha_8_label.grid(row=4, column=0, padx=10, pady=5, sticky="w")
        self.malha_8_entry = ttk.Entry(input_frame)
        self.malha_8_entry.grid(row=4, column=1, padx=10, pady=5)
        self.malha_8_entry.bind("<Return>", lambda event: self.focus_next(event, self.malha_100_entry))

        malha_100_label = tk.Label(input_frame, text="Valor da Malha 100:")
        malha_100_label.grid(row=5, column=0, padx=10, pady=5, sticky="w")
        self.malha_100_entry = ttk.Entry(input_frame)
        self.malha_100_entry.grid(row=5, column=1, padx=10, pady=5)
        self.malha_100_entry.bind("<Return>", lambda event: self.focus_next(event, self.fundo_entry))

        fundo_label = tk.Label(input_frame, text="Valor do Fundo:")
        fundo_label.grid(row=6, column=0, padx=10, pady=5, sticky="w")
        self.fundo_entry = ttk.Entry(input_frame)
        self.fundo_entry.grid(row=6, column=1, padx=10, pady=5)
        self.fundo_entry.bind("<Return>", lambda event: self.calculate_and_display())

        # Botão para calcular e exibir resultados
        calculate_button = ttk.Button(self.root, text="Calcular e Exibir", command=self.calculate_and_display)
        calculate_button.pack(pady=10)

        # Frame para exibir os resultados
        self.result_frame = ttk.LabelFrame(self.root, text="Resultados")
        self.result_frame.pack(pady=20, padx=10, fill="both")

    def focus_next(self, event, next_entry):
        next_entry.focus_set()

    def autocomplete_date(self, event):
        current_text = self.data_entry.get()
        if len(current_text) == 2 or len(current_text) == 5:
            self.data_entry.insert(tk.END, '/')

    def autocomplete_time(self, event):
        current_text = self.hora_entry.get()
        if len(current_text) == 2:
            self.hora_entry.insert(tk.END, ':')

    def calculate_and_display(self):
        try:
            nome = self.nome_entry.get()
            data = self.data_entry.get()
            hora = self.hora_entry.get()
            malha_7 = float(self.malha_7_entry.get())
            malha_8 = float(self.malha_8_entry.get())
            malha_100 = float(self.malha_100_entry.get())
            fundo = float(self.fundo_entry.get())

            # Calcula a soma total
            soma_total = malha_7 + malha_8 + malha_100 + fundo

            # Calcula os percentuais
            percentual_malha_7 = (malha_7 * 100) / soma_total
            percentual_malha_8 = (malha_8 * 100) / soma_total
            percentual_malha_100 = (malha_100 * 100) / soma_total
            percentual_fundo = (fundo * 100) / soma_total

            # Exibe os resultados
            resultados_text = (
                f"Valores digitados:\n"
                f"Malha 7: {malha_7}\n"
                f"Malha 8: {malha_8}\n"
                f"Malha 100: {malha_100}\n"
                f"Fundo: {fundo}\n\n"
                f"Soma total: {soma_total}\n\n"
                f"Percentuais:\n"
                f"Malha 7 em porcentagem: {percentual_malha_7:.1f}%\n"
                f"Malha 8 em porcentagem: {percentual_malha_8:.1f}%\n"
                f"Malha 100 em porcentagem: {percentual_malha_100:.1f}%\n"
                f"Fundo em porcentagem: {percentual_fundo:.1f}%"
            )

            # Atualiza o frame de resultados
            for widget in self.result_frame.winfo_children():
                widget.destroy()

            result_label = tk.Label(self.result_frame, text=resultados_text, justify="left")
            result_label.pack()

            # Salvar os dados em um arquivo Excel
            data_dict = {
                'Nome': [nome],
                'Data': [data],
                'Hora': [hora],
                'Malha 7': [malha_7],
                'Malha 8': [malha_8],
                'Malha 100': [malha_100],
                'Fundo': [fundo],
                'Soma Total': [soma_total],
                'Percentual Malha 7': [percentual_malha_7],
                'Percentual Malha 8': [percentual_malha_8],
                'Percentual Malha 100': [percentual_malha_100],
                'Percentual Fundo': [percentual_fundo]
            }

            df = pd.DataFrame(data_dict)

            # Verifica se o arquivo já existe
            if os.path.exists('dados.xlsx'):
                df_existing = pd.read_excel('dados.xlsx')
                df = pd.concat([df_existing, df], ignore_index=True)

            # Salva no Excel
            df.to_excel('dados.xlsx', index=False)

            # Gerar gráfico comparativo
            self.generate_comparison_chart(df)

        except ValueError as e:
            messagebox.showerror("Erro", f"Erro na entrada de dados: {str(e)}")

    def generate_comparison_chart(self, df):
        try:
            # Gera o gráfico comparativo
            fig, ax = plt.subplots(figsize=(10, 6))

            # Concatena data e hora para criar uma coluna datetime completa
            df['DataHora'] = pd.to_datetime(df['Data'] + ' ' + df['Hora'], format='%d/%m/%Y %H:%M')

            # Plota os dados
            sns.barplot(x='DataHora', y='Malha 7', data=df, ax=ax, label='Malha 7', color='blue')
            sns.barplot(x='DataHora', y='Malha 8', data=df, ax=ax, label='Malha 8', color='orange')
            sns.barplot(x='DataHora', y='Malha 100', data=df, ax=ax, label='Malha 100', color='green')
            sns.barplot(x='DataHora', y='Fundo', data=df, ax=ax, label='Fundo', color='red')

            ax.set_title('Comparativo de Valores por Data e Hora', fontsize=16, fontweight='bold')
            ax.set_xlabel('Data e Hora', fontsize=12)
            ax.set_ylabel('Valores', fontsize=12)
            ax.grid(True)
            plt.xticks(rotation=45)
            ax.legend()

            # Adiciona data e hora como texto no topo do gráfico
            date_str = df['Data'].iloc[-1]  # Pega a última data inserida
            time_str = df['Hora'].iloc[-1]  # Pega a última hora inserida
            text_str = f'Última atualização: {date_str} {time_str}'
            plt.text(0.02, 0.95, text_str, transform=ax.transAxes, fontsize=12, verticalalignment='top')

            # Salvar o gráfico como PNG na pasta 'graficos'
            if not os.path.exists('graficos'):
                os.makedirs('graficos')

            filename = os.path.join('graficos', 'grafico_comparativo.png')
            plt.savefig(filename)

            # Exibir o gráfico em uma nova janela
            self.show_chart_in_window(fig)

        except Exception as e:
            messagebox.showerror("Erro", f"Erro ao gerar o gráfico: {str(e)}")

    def show_chart_in_window(self, fig):
        graph_window = tk.Toplevel(self.root)
        graph_window.title('Gráfico Comparativo')
        graph_window.geometry('800x600')

        canvas = FigureCanvasTkAgg(fig, master=graph_window)
        canvas.draw()
        canvas.get_tk_widget().pack(fill=tk.BOTH, expand=True)


if __name__ == "__main__":
    root = tk.Tk()
    app = DashboardApp(root)
    root.mainloop()

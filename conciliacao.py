from tkinter import *
from tkinter import ttk
import tkinter.messagebox
import pandas as pd
from datetime import datetime
from dateutil.relativedelta import relativedelta
import openpyxl
import os
import numpy as np


class Conciliacao:

    def __init__(self, janela):
        self.janela = janela
        self.janela.title('Conciliações')
        self.janela.geometry('1000x600')

        self.frame1 = Frame(self.janela, width=800, height=400, bg='white', bd=5, relief=RIDGE).\
            grid(padx=100, pady=100)
        Label(self.frame1, text='Login', font=('Impact', 35, 'bold'), fg='#6162FF', bg='white').\
            place(x=440, y=130)
        Label(self.frame1, text='Usuario', font=('Goudy old style', 15, 'bold'), fg='grey', bg='white'). \
            place(x=400, y=220)

        self.usuario = ttk.Combobox(self.frame1, font=('arial', 14, 'bold'), width=15)
        self.usuario['values'] = ('Leandro Peixoto', 'Mariclea Martini',
                                  'Michele Bernardino', 'Paulo França')
        # self.usuario.current(0)
        self.usuario.place(x=400, y=250)

        self.btn_entrar = Button(self.frame1, text='Entrar', font=('Goudy old style', 15, 'bold'), width=10,
                                 bd=0, bg='#6162FF', fg='white', command=self.tela_inicial)
        self.btn_entrar.place(x=440, y=350)

    def tela_inicial(self):
        self.janela.withdraw()
        self.inicio = Toplevel()
        self.inicio.geometry('1000x600')

        self.tela_frame = Frame(self.inicio, width=800, height=400, bg='white')
        self.tela_frame.place(x=100, y=100)
        Label(self.tela_frame, text='Selecione a Competência', font=('arial', 16, 'bold')). \
            place(x=270, y=50)
        lista = []
        for i in range(12):
            mes = datetime.today()
            data_limite = mes - relativedelta(months=i)
            lista.append(data_limite.strftime('%m/%Y'))

        fonte = ('arial', 14)
        self.competencia = ttk.Combobox(self.tela_frame, font=('arial', 16, 'bold'), width=15)
        self.competencia['values'] = (lista)
        self.inicio.option_add('*TCombobox*Listbox.font', fonte)
        # self.competencia.current(0)
        self.competencia.place(x=300, y=150)
        self.verifica = Button(self.tela_frame, text='Gerar Relatório', bd=5, font=('arial', 16, 'bold'), command=self.validacao)
        self.verifica.place(x=310, y=250)

    def validacao(self):
        pasta1 = os.listdir('G:\GECOT\CONCILIAÇÕES CONTÁBEIS\CONCILIAÇÕES_2021\\11.2021')

        lista = [[], [], [], [], []]

        for i in pasta1[::-1]:
            if i.startswith('~') == True:
                pasta1.remove(i)

        for i in pasta1:
            wb = openpyxl.load_workbook('G:\GECOT\CONCILIAÇÕES CONTÁBEIS\CONCILIAÇÕES_2021\\11.2021\\' + i,
                                        read_only=True)
            sheets = wb.sheetnames
            ws = wb[sheets[0]]
            try:
                conta = ws['A2'].value.split()
            except:
                conta = ['', '']
            valor_deb = ws['C5'].value
            valor_cred = ws['D5'].value
            data = ws['A5'].value.strftime('%m/%Y')
            lista[0].append(conta[1])
            lista[1].append(data)
            lista[2].append(valor_deb)
            lista[3].append(valor_cred)

            wb.close()

        data = pd.DataFrame(lista).T

        data.columns = ['Conta', 'Data', 'Debito', 'Credito', 'Saldo']

        dados = pd.read_excel('balteste.xlsx')
        dados = pd.DataFrame(dados)

        apoio = pd.read_excel('contas.xlsx')
        apoio = pd.DataFrame(apoio)

        for index1, row1 in data.iterrows():
            for index, row in dados.iterrows():
                if row1['Conta'] == row['Conta CSPE']:
                    # data.insert(3, 'Saldo', '')
                    data['Saldo'].loc[index1] = dados.loc[index, 'Saldo Acumulado']

        data[['Debito', 'Credito', 'Saldo']] = data[['Debito', 'Credito', 'Saldo']].apply(pd.to_numeric)

        data.fillna(0, inplace=True)
        data = data.round(2)
        data['Resultado'] = data['Debito'] - data['Credito'] - data['Saldo']

        # data['Resultado'] = data.apply(lambda x: x['Debito'] - x['Saldo'], axis=1)
        data = pd.merge(data, apoio[['Conta', 'Usuario']], on=['Conta'], how='left')

        data['Status'] = np.where(data['Resultado'] != 0, 'Diferença de Valor', 'OK')

        # data = data.loc[data['Usuario'] == self.usuario.get()]


        data.to_excel('teste.xlsx', index=False)

        self.data = data
        tkinter.messagebox.showinfo('', 'Arquivo Validado com Sucesso')
        self.relatorio()

    def relatorio(self):
        self.inicio.withdraw()
        self.relat = Toplevel()
        self.relat.geometry('1000x600')

        self.val_frame = Frame(self.relat, width=800, height=400, bg='white')
        self.val_frame.place(x=100, y=100)

        estilo = ttk.Style()
        estilo.theme_use('default')
        estilo.configure('Treeview', background='#D3D3D3', foreground='black', rowheight=25,
                         fieldbackground='#D3D3D3')
        estilo.map('Treeview', background=[('selected', '#347083')])

        # Treeview frame
        tree_frame = Frame(self.val_frame)
        tree_frame.pack(pady=50)
        # Barra rolagem
        tree_scroll = Scrollbar(tree_frame)
        tree_scroll.pack(side=RIGHT, fill=Y)
        # Criar Treeview
        nf_tree = ttk.Treeview(tree_frame, yscrollcommand=tree_scroll.set, selectmode='extended')
        nf_tree.pack(side=LEFT)
        # Configurar Barra Rolagem
        tree_scroll.config(command=nf_tree.yview)
        # Definir colunas
        colunas2 = ['Conta', 'Data', 'Debito', 'Credito', 'Saldo', 'Resultado', 'Usuario', 'Status']
        nf_tree['columns'] = colunas2
        # formatar colunas
        nf_tree.column('Conta', width=80)
        nf_tree.column('Data', width=60)
        nf_tree.column('Debito', width=100)
        nf_tree.column('Credito', width=100)
        nf_tree.column('Saldo', width=100)
        nf_tree.column('Resultado', width=100)
        nf_tree.column('Usuario', width=100)
        nf_tree.column('Status', width=140)

        # formatar títulos
        nf_tree.heading('Conta', text='Conta', anchor=W)
        nf_tree.heading('Data', text='Data', anchor=W)
        nf_tree.heading('Debito', text='Debito', anchor=W)
        nf_tree.heading('Credito', text='Credito', anchor=W)
        nf_tree.heading('Saldo', text='Saldo', anchor=W)
        nf_tree.heading('Resultado', text='Resultado', anchor=W)
        nf_tree.heading('Usuario', text='Usuario', anchor=W)
        nf_tree.heading('Status', text='Status', anchor=W)

        nf_tree['show'] = 'headings'

        # inserir dados do banco no treeview
        def inserir_tree(lista):
            nf_tree.delete(*nf_tree.get_children())
            contagem = 0
            for index, row in lista.iterrows():  # loop para inserir cores diferentes nas linhas
                if row[7] == 'OK':
                    nf_tree.insert(parent='', index='end', text='', iid=contagem,
                                   values=(row[0], row[1], row[2], row[3], row[4], row[5], row[6], row[7]), tags=('evenrow',))
                else:
                    nf_tree.insert(parent='', index='end', text='', iid=contagem,
                                   values=(row[0], row[1], row[2], row[3], row[4], row[5], row[6], row[7]), tags=('oddrow',))
                contagem += 1

        nf_tree.tag_configure('oddrow', background='light coral')
        nf_tree.tag_configure('evenrow', background='lightgreen')
        inserir_tree(self.data)

        def NotasInfo2(ev):
            verinfo2 = nf_tree.focus()
            dados2 = nf_tree.item(verinfo2)
            row = dados2['values']
            os.startfile('G:\GECOT\CONCILIAÇÕES CONTÁBEIS\CONCILIAÇÕES_2021\\11.2021' +
                         '\\' + 'Conta '+ row[0].replace('.', '') + '.xlsx')

        nf_tree.bind('<Double-Button>', NotasInfo2)

if __name__=='__main__':
    janela = Tk()
    aplicacao = Conciliacao(janela)
    janela.mainloop()

import pandas as pd
import os
import tkinter as tk
from tkinter import filedialog,ttk,messagebox

def convertNumeric(value):
    try:
        return pd.to_numeric(value)
    except ValueError:
        return value

def Ler_reatorios():
    file = filedialog.askopenfilename(title='Escolha o arquivo',filetypes=[('CSV Files','*.csv')])
    if file:
        df = pd.read_csv(file, delimiter=';')
        calcularValor(df)
        #print(df)
    else:
        messagebox.showerror("Erro","Escolha um arquivo em formato .CSV")
        return

def calcularValor(arquivo):
    df = arquivo

    # Declarando as variáveis das contas
    global ReceitaLiquida
    global DespesasGerais
    global ProvisaoPerdas
    global OutrasReceitas
    global conta_107
    global conta_135
    global conta_81
    global CustoServico
    global LucroBruto
    global ResultadoSemReceitas
    global ResultadoAntesImposto
    global LucroLiquido
    global RecFinanLiquidas
    global Depreciacao
    global conta_328
    global conta_378

    # Colunas
    ind_coluna = 0 # Coluna 0
    vlr_coluna = 6 # Coluna 6

    # Calculando a RECEITA LIQUIDA #
    conta_141 = df[df.iloc[:, ind_coluna] == '141'].iloc[:,vlr_coluna].values
    conta_214 = df[df.iloc[:, ind_coluna] == '214'].iloc[:,vlr_coluna].values

    conta_141 = float(conta_141[0].replace('.','').replace(',','.'))
    conta_214 = float(conta_214[0].replace('.','').replace(',','.'))

    ReceitaLiquida = conta_141 + conta_214

    # Pegando o valor do CUSTOS DOS SERVIÇOS PRESTADOS E MERCADORIAS VENDIAS #
    CustoServico = df[df.iloc[:, ind_coluna] == '219'].iloc[:,vlr_coluna].values

    # Calculando o LUCRO BRUTO
    if CustoServico.size > 0:
        float(CustoServico[0].replace('.','').replace(',','.'))
        LucroBruto = ReceitaLiquida - CustoServico
    else:
        LucroBruto = ReceitaLiquida

    # Calculando as DESPESAS GERAIS E ADMINISTRATIVAS #
    conta_125= df[df.iloc[:, ind_coluna] == '125'].iloc[:,vlr_coluna].values
    conta_126 = df[df.iloc[:, ind_coluna] == '126'].iloc[:,vlr_coluna].values
    conta_127 = df[df.iloc[:, ind_coluna] == '127'].iloc[:,vlr_coluna].values

    conta_125 = float(conta_125[0].replace('.','').replace(',','.'))
    conta_126 = float(conta_126[0].replace('.','').replace(',','.'))
    conta_127 = float(conta_127[0].replace('.','').replace(',','.'))

    DespesasGerais = conta_125 + conta_126 + conta_127

    # Calculando a PROVISÃO PARA PERDA ESPERADA DOS SERVIÇOS FATURADOS #
    conta_325 = df[df.iloc[:, ind_coluna] == '325'].iloc[:,vlr_coluna].values
    conta_171 = df[df.iloc[:, ind_coluna] == '171'].iloc[:,vlr_coluna].values

    conta_325 = float(conta_325[0].replace('.','').replace(',','.'))
    conta_171 = float(conta_171[0].replace('.','').replace(',','.'))

    ProvisaoPerdas = conta_325 + conta_171

    # Calculando OUTRAS RECEITAS E DESP. OP. #
    conta_376 = df[df.iloc[:, ind_coluna] == '376'].iloc[:,vlr_coluna].values
    conta_405 = df[df.iloc[:, ind_coluna] == '405'].iloc[:,vlr_coluna].values

    conta_376 = float(conta_376[0].replace('.','').replace(',','.'))
    conta_405 = float(conta_405[0].replace('.','').replace(',','.'))
    
    OutrasReceitas = conta_376 + conta_405

    # Calculando o RESULTADO ANTES DAS RECEITAS(DESPESAS FINANCEIRAS), RESULTADO DE EQUIVALÊNCIA PATRIMONIAL E IMPOSTOS #
    ResultadoSemReceitas = LucroBruto + DespesasGerais + ProvisaoPerdas + OutrasReceitas

    # Pegando as RECEITAS FINANCEIRAS # 
    conta_107 = df[df.iloc[:, ind_coluna] == '107'].iloc[:,vlr_coluna].values
    conta_107 = float(conta_107[0].replace('.','').replace(',','.'))

    # Pegando as DESPESAS FINANCEIRAS #
    conta_135 = df[df.iloc[:, ind_coluna] == '135'].iloc[:,vlr_coluna].values
    conta_135 = float(conta_135[0].replace('.','').replace(',','.'))

    # Calculando as RECEITAS(DESPESAS) FINANCEIRAS, LÍQUIDAS
    RecFinanLiquidas = conta_107 + conta_135

    # Calculando o RESULTADO ANTES DO IMPOSTO DE RENDA E DA CONTRIBUIÇÃO SOCIAL #
    ResultadoAntesImposto = ResultadoSemReceitas + conta_107 + conta_135

    # Pegando o IR/CSLL #
    conta_81= df[df.iloc[:, ind_coluna] == '81'].iloc[:,vlr_coluna].values
    conta_81 = float(conta_81[0].replace('.','').replace(',','.'))

    # Calculando o LUCRO LÍQUIDO #
    LucroLiquido = ResultadoAntesImposto + conta_81

    # Pegando a DEPRECIACAO ACUM.OPERACIONAL
    conta_328 = df[df.iloc[:,ind_coluna] == '328'].iloc[:,vlr_coluna].values
    conta_328 = conta_328[0].replace('-','')
    conta_328 = float(conta_328.replace('.','').replace(',','.'))

    # Pegando a VENDA ATIVO IMOBILIZADO
    conta_378 = df[df.iloc[:,ind_coluna] == '378'].iloc[:,vlr_coluna].values
    conta_378 = float(conta_378[0].replace('.','').replace(',','.'))

def gerarRelatorio():
    global comboboxMes
    global comboboxAno

    # Coluna 1
    selected_month = comboboxMes.get()
    selected_year = comboboxAno.get()

    palavras = ['Cindapa do Brasil',f'{selected_month}/{selected_year}','Receita Líquida dos serviços prestados e mercadorias vendidas',
                '(-) Custos os serviços prestados e mercadorias vendidas', '(=) Lucro Bruto', '(-) Despesas gerais e administrativas',
                '(-) Provisão para perda esperada dos serviços faturados','(-) Outras receitas e desp. Op.', '(=) Resultado antes das receitas(despesas financeiras), resultado de equivalência patrimonial e impostos',
                '(+) Receitas financeiras', '(-) Despesas financeiras', 'Receitas(despesas) financeiras, líquidas', 'Resultado antes do imposto de renda e da contribuição social',
                '(-) IR/CSLL', '(=) Lucro líquido','','','EBITDA','Lucro Líquido','Resultado Financeiro','Depreciação','Impostos sobre o lucro','Venda ativo imobilizado','EBITDA']
    
    # Coluna 2
    # Puxando as variáveis das contas com os valores
    global ReceitaLiquida,LucroBruto,DespesasGerais,ProvisaoPerdas,OutrasReceitas,conta_107,conta_135,conta_81,CustoServico,ResultadoAntesImposto,ResultadoSemReceitas,LucroLiquido,RecFinanLiquidas

    NewRecFinanLiquidas = float(str(RecFinanLiquidas).replace('-',''))
    NewConta81 = float(str(conta_81).replace('-',''))

    ValorEbitda = LucroLiquido + NewRecFinanLiquidas + conta_328 + NewConta81 - conta_378
    valores = ['','',ReceitaLiquida,CustoServico,LucroBruto,DespesasGerais,ProvisaoPerdas,OutrasReceitas,ResultadoSemReceitas,conta_107,conta_135,RecFinanLiquidas,ResultadoAntesImposto,conta_81,LucroLiquido,'','','Valor',LucroLiquido,NewRecFinanLiquidas,conta_328,NewConta81,conta_378,ValorEbitda]

    # Coluna 3
    PorcentoLucroLiquido = (LucroLiquido/ReceitaLiquida)*100
    PorcentoResultFinanceiro = (NewRecFinanLiquidas/ReceitaLiquida)*100
    PorcentoDeprecia = (conta_328/ReceitaLiquida)*100
    PorcentoImpostosLucro = (NewConta81/ReceitaLiquida)*100
    PorcentoAtivoImo = (conta_378/ReceitaLiquida)*100
    TotalPorcento = PorcentoLucroLiquido+PorcentoResultFinanceiro+PorcentoDeprecia+PorcentoImpostosLucro-PorcentoAtivoImo

    coluna3 = ['','','','','','','','','','','','','','','','','','%',f'{PorcentoLucroLiquido:.2f}%',f'{PorcentoResultFinanceiro:.2f}%',f'{PorcentoDeprecia:.2f}%',f'{PorcentoImpostosLucro:.2f}%',f'{PorcentoAtivoImo:.2f}%',f'{TotalPorcento:.2f}%']

    # Aplicando os valores e palavras nas colunas
    df_final = pd.DataFrame({'Coluna1':palavras,'Coluna2':valores,'Coluna3':coluna3})

    #df_final['Coluna2'] = df_final['Coluna2'].replace(",",".")
    #df_final['Coluna2'] = pd.to_numeric(df_final['Coluna2'], errors='coerce')

    df_final['Coluna2'] = df_final['Coluna2'].apply(convertNumeric)
    downloads_folder = os.path.join(os.path.expanduser("~"),"Downloads")
    df_final.to_excel(os.path.join(downloads_folder,"Resumo_Ebtida.xlsx"),index=False,header=False,float_format="%.2f")

def janela():
    window = tk.Tk()
    window.title("OnLine Contabilidade - Ebtida")
    window.geometry("570x220")
    window.iconbitmap('logoOnline.ico')

    # Fixa o tamanho da janela
    window.resizable(False,False)

    meses_num = ['Janeiro','Fevereiro','Março','Abril','Maio','Junho','Julho','Agosto','Setembro','Outubro','Novembro','Dezembro']

    anos = ['2024','2025','2026','2027','2028','2029','2030']

    # Configurar colunas e linhas para centralizar os widgets
    window.columnconfigure(0, weight=1)
    window.columnconfigure(1, weight=1)
    window.columnconfigure(2, weight=1)
    window.rowconfigure(0, weight=1)
    window.rowconfigure(1, weight=1)
    window.rowconfigure(2, weight=1)
    window.rowconfigure(3, weight=1)
    window.rowconfigure(4, weight=1)

    # Escolha do mês #
    labelTexto = tk.Label(window, text="Escolha o mês:")
    labelTexto.grid(row=1, column=1, sticky='e')

    global comboboxMes
    comboboxMes = ttk.Combobox(window, values=meses_num)
    comboboxMes.grid(row=1, column=2, sticky='w')

    # Escolha do ano #
    labelAno = tk.Label(window,text='Escolha o ano:')
    labelAno.grid(row=2,column=1,sticky='e')

    global comboboxAno
    comboboxAno = ttk.Combobox(window,values=anos)
    comboboxAno.grid(row=2,column=2,sticky='w')

    # Botao para escolher o arquivo #
    btn1 = tk.Button(window, text='Escolha o arquivo', width=70, height=2, command=Ler_reatorios)
    btn1.grid(row=3, column=1, columnspan=2, pady=5)

    # Botao para gerar o relatorio #
    btn2 = tk.Button(window, text='Gerar relatório Ebtida', width=70, height=2, command=gerarRelatorio)
    btn2.grid(row=4, column=1, columnspan=2, pady=5)

    # Botao para fechar o programa #
    btn3 = tk.Button(window, text='Sair', width=70, height=2, command=window.destroy)
    btn3.grid(row=5, column=1, columnspan=2, pady=5)

    window.mainloop()
janela()
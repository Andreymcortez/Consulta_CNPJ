import pandas as pd
import datetime, time, requests, json, tkinter, warnings
from tkinter import *
from tkinter import filedialog, messagebox
from tkinter.ttk import *

warnings.simplefilter(action='ignore', category=FutureWarning)
# Cria janela
janela = Tk()

def consulta_cnpj():

    # Criar janela de seleção
    root = Tk()
    
    # Não aparecer a janela root
    root.wm_withdraw()
    
    # Criar função do botão de selecionar arquivo
    root.filename =  filedialog.askopenfilename(initialdir = '/',title = 'Select file',
                                                filetypes = (('Arquivos Excel','*.xlsx'),
                                                             ('all files','*.*'),
                                                             ('Arquivos CSV','*.csv')))
    
    # Carregar o arquivo selecionado em um dataframe
    cnpjs_df = pd.read_csv(root.filename, encoding='latin1', sep=';')
    
    
    # Iniciar cronometro de tempo de execução do script
    inicio = time.perf_counter()
    ti = datetime.datetime.now().strftime('%d-%m-%Y %H:%M:%S')
    messagebox.showinfo(title='Aviso',message=f'\nIniciado em: {ti}\n',parent=janela)
    
    base = pd.DataFrame(columns=[
                                    'cnpj',
                                    'razao_social',
                                    'situacao_cadastral',
                                    'data_situacao_cadastral',
                                    'descricao_motivo_situacao_cadastral',
                                    'data_inicio_atividade']
                        )
    
    # Iniciar o loop de pesquisa de CNPJ
    for cnpj in cnpjs_df['CNPJ']:
        df_cnpj = pegarCNPJ({'cnpj':cnpj})
        jCNPJ(df_cnpj)
        if df_cnpj != None:
            linhas = {
                        'cnpj': cnpj,
                        'razao_social': df_cnpj['razao_social'],
                        'situacao_cadastral': df_cnpj['descricao_situacao_cadastral'],
                        'motivo_situacao_cadastral': df_cnpj['descricao_motivo_situacao_cadastral'],
                        'data_situacao_cadastral': df_cnpj['data_situacao_cadastral'],
                        'descricao_motivo_situacao_cadastral': df_cnpj['descricao_motivo_situacao_cadastral'],
                        'data_inicio_atividade': df_cnpj['data_inicio_atividade']
                        }
        else:
            linhas = {'cnpj': cnpj}
        base = base.append(linhas, ignore_index=True)
        teste = base['teste'] = base['cnpj'] + ' - Situação Cadastral: ' + base['situacao_cadastral']
        print('*' *50)
        print(teste.to_string())
        print('*' *50)
        base.drop('teste', inplace=True, axis=1)
        
        # Nome do arquivo para salvar a consulta
        base.to_excel('Consulta_CNPJ(Oceana).xlsx',index=False)
    progress.pack(pady = 10)
        
    # Finalizar o contador
    final = time.perf_counter()
    tf = datetime.datetime.now().strftime('%d-%m-%Y %H:%M:%S')
    if round(final - inicio, 1) > 60:
        total = round((final - inicio) / 60)
        if total < 2:
            messagebox.showinfo(title='Aviso',message=f'Finalizado em: {ti}\n\nTempo de execução: {total} Minuto',parent=janela)
        else:
            messagebox.showinfo(title='Aviso',message=f'Finalizado em: {ti}\n\nTempo de execução: {total} Minutos',parent=janela)
    else:
        total = round(final - inicio)
        if total >= 2:
            messagebox.showinfo(title='Aviso',message=f'Finalizado em: {ti}\n\nTempo de execução: {total} Segundos',parent=janela)
        else:
            messagebox.showinfo(title='Aviso',message=f'Finalizado em: {ti}\n\nTempo de execução: {total} Segundo',parent=janela)

# Fazer as requisições para a API do Minha Receita. Caso bem sucedido, retorna o conteúdo em formato JSON
def pegarCNPJ(cnpj):
    APIReceita = 'https://minhareceita.org/'
    r = requests.post(APIReceita, data=cnpj, timeout=None)
    if r.status_code == 200:
        return json.loads(r.content)

# Função de visualização de json
def jCNPJ(cnpj):
    info = json.dumps(cnpj, sort_keys=True, indent=4, ensure_ascii=False)

# Barra de progresso
progress = Progressbar(janela, orient = HORIZONTAL,length = 100, mode = 'determinate', maximum=100)
def bar():
    for i in range(10):# range(len(cnpjs_df)):
        progress['value'] = i
        janela.update_idletasks()
        time.sleep(0.1)

# Não permitir ajustar o tamanho da janela
janela.resizable(width=False, height=False)

# Colocar imagem de funda na janela
bg = PhotoImage(file = 'bg.png')
canvas = Canvas(janela, width = 450, height = 300)
canvas.pack(fill = 'both', expand = True) 
canvas.create_image( 0, 0, image = bg, anchor = 'nw')
canvas.create_text( 295, 130, text = 'Selecione o arquivo com os CNPJs', fill='white')
canvas.create_text( 420, 570, text = 'Consulta de CNPJ - By Andreydson Cortez', fill='white')

def fechar_janela():
    janela.destroy()

# Criar botões
botao_selecionar = Button(janela,  text='Selecionar arquivo', command=consulta_cnpj) 
botao_canvas = canvas.create_window( 235, 150, anchor = 'nw', window = botao_selecionar)

botao_fechar = Button(janela,  text='Fechar janela', command=fechar_janela)
botao_canvas = canvas.create_window( 245, 410, anchor = 'nw', window = botao_fechar)

barra_progresso = Progressbar(janela, orient= HORIZONTAL, length = 100, mode='indeterminate')

janela.title('Consulta CNPJ - Andrey Cortez')
janela.geometry('600x600')
janela.mainloop()

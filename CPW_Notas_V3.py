from fpdf import FPDF;
import datetime, os;
from datetime import date;
import mysql.connector;
from mysql.connector.errors import Error;
from mysql.connector import errorcode;
import glob;
import ctypes
import sys

nomeservidor=("localhost")
porta=3306
usuario='user'
key='senha'
banco='banco'

def Sair():
    sys.exit()

def Procedimento():
    ctypes.windll.user32.MessageBoxW('AGUARDE O PROCEDIMENTO SER FINALIZADO!', 'COMPUWAY - NOTAS', 0)
    try: #try conexão com banco de dados
        db_connection = mysql.connector.connect(host=nomeservidor, port=porta, user=usuario, passwd=key, database=banco)
        sql = ("SELECT matricula_aluno, nome_aluno, sobrenome_aluno, nome_curso, porc_teste_concluido, carga_horaria, data_teste_concluido from teste_concluido INNER JOIN aluno ON aluno.id_aluno = teste_concluido.id_aluno INNER JOIN curso ON curso.cod_curso = teste_concluido.cod_curso WHERE aula_teste_concluido='final' ORDER BY data_teste_concluido DESC LIMIT 1000;")
        cursor = db_connection.cursor()
        cursor.execute(sql)
        dados = cursor.fetchall()

        for i in range(len(dados)):
            if dados[i][4] == 100:
                dados[i] = dados[i][0], dados[i][1], dados[i][2], dados[i][3] , 10, dados[i][5], dados[i][6]
            else:
                notaaluno=((dados[i][4])/10)
                dados[i] = dados[i][0], dados[i][1], dados[i][2], dados[i][3] , notaaluno, dados[i][5], dados[i][6]  

        cabecalho=['MATRICULA','NOME','SOBRENOME','CURSO','NOTA','CARGA (h)','DATA']
        dados.insert(0,cabecalho) 

        cursor.close()
        db_connection.commit()
        db_connection.close()

        try:
            pdf = FPDF(orientation = 'L', unit = 'mm', format='A4')
            pdf.add_page()
            pdf.set_font('Times','',24.0)  
            epw = pdf.w - 2*pdf.l_margin   
            col_width = epw/7
            th = pdf.font_size                     
            pdf.set_text_color(255, 102, 0)
            pdf.ln(2*th)                           
            lh_list = [] 
            use_default_height = 3 #flag
            pdf.cell(epw, 0.0, 'COMPUWAY - NOTAS ALUNOS', align='C')                      
            pdf.ln(2*th)                         
            pdf.set_text_color(0, 0, 0)
            pdf.set_font('Times','',8.0) 
            for row in dados:
                for datum in row: 
                    pdf.cell(col_width, th, str(datum), border=1)
                pdf.ln(th)
            PathPadrão = os.path.exists('C:\COMPUWAY')
            if not PathPadrão: # Se não encontrado o caminho C:\COMPUWAY é CRIADO pasta PADRÃO
                os.makedirs('C:\COMPUWAY')
            arquivoNome="C:\COMPUWAY\ALUNOS_NOTAS.pdf"
            pdf.output(arquivoNome) 
            imprimir = ctypes.windll.user32.MessageBoxW(0,'Deseja imprimir arquivo Notas?', 'COMPUWAY - NOTAS', 4)
            if imprimir == 6:
                list_of_files = glob.glob('C:\COMPUWAY\*.pdf')
                latest_file = max(list_of_files, key=os.path.getctime)
                os.startfile(latest_file)
            else:
                Sair()
        except ValueError:
            ctypes.windll.user32.MessageBoxW(0,'NÃO FOI POSSÍVEL EXECUTAR O PROGRAMA!\nERRO: 9999', 'ERRO!', 16)
       
    except mysql.connector.Error as err:
                                if err.errno == errorcode.ER_ACCESS_DENIED_ERROR:
                                    ctypes.windll.user32.MessageBoxW(0,'NÃO FOI POSSÍVEL EXECUTAR O PROGRAMA!\nSEM PERMISSÃO\nERRO: 9901', 'ERRO!', 16)
                                    Sair()
                                elif err.errno == errorcode.ER_BAD_DB_ERROR:
                                    ctypes.windll.user32.MessageBoxW(0,'NÃO FOI POSSÍVEL EXECUTAR O PROGRAMA!\SERVIDOR NÃO EXISTE!\nERRO: 9902', 'ERRO!', 16)
                                    Sair()
                                else:
                                    ctypes.windll.user32.MessageBoxW(0,'NÃO FOI POSSÍVEL EXECUTAR O PROGRAMA!\n\nERRO: 9900', 'ERRO!', 16)
                                    Sair()

PathGerenciador = os.path.exists('C:\Program Files (x86)\Cursos Compuway\Gerenciador de Alunos')
if not PathGerenciador: # Verifica se diretorio do Gerenciador de Alunos NÃO é ENCONTRADO
    ctypes.windll.user32.MessageBoxW(0,'A Máquina NÃO possui GERENCIADOR DE ALUNOS\nInicie o programa na máquina servidor!\nO Programa será encerrado...', 'ERRO!', 16)
    Sair()
else:
    respostaentrada = ctypes.windll.user32.MessageBoxW(0,'Deseja realizar o procedimento?', 'COMPUWAY - NOTAS', 1)
    if respostaentrada==1:
        Procedimento()
    else:
        Sair()
from datetime import date
import mysql.connector;
from mysql.connector.errors import Error;
from mysql.connector import errorcode;
from fpdf import FPDF;
import os, datetime, shutil,subprocess;
from tkinter import Tk, Label,Button, Entry, messagebox, ttk, filedialog, font, LabelFrame,PhotoImage;
import glob;

Janela = Tk()
Janela.geometry("320x200")          # TAMANHO DA JANELA
Janela.title("COMPUWAY - BACKUP")   # NOME DA JANELA
Janela.configure(bg='#414141')      # COR DA JANELA
Janela.wait_visibility(Janela)
Janela.wm_attributes('-alpha',0.9)
Janela.resizable(False, False)      # DESATIVA REAJUSTAR A JANELA
Janela.overrideredirect(1)          # REMOVE HUD DA JANELA
Janela.attributes('-topmost',True) # SEMPRE NA FRENTE DE OUTROS PROGRAMAS

# CENTRALIZAR PROGRAMA
windowWidth = Janela.winfo_reqwidth() 
windowHeight = Janela.winfo_reqheight()
positionRight = int(Janela.winfo_screenwidth()/2 - windowWidth/2)
positionDown = int(Janela.winfo_screenheight()/2 - windowHeight/2)
Janela.geometry("+{}+{}".format(positionRight, positionDown))

titleFont = font.Font(family='Helvetica', size=16, weight='bold')
btnFont = font.Font(family='montserrat', size=12, weight='bold')
lblFont = font.Font(family='montserrat', size=8, weight='bold')

def Sair(): # Função para fechar Janela
        Janela.destroy()  

def enter_pressed(event): # Função para verificar se tecla Enter é pressiona ao foco nos seguintes intes: entryMatricula e entryLogin
   Nota() 

def Nota():
    try:
        PathGerenciador = os.path.exists('C:\Program Files (x86)\Cursos Compuway\Gerenciador de Alunos')
        if not PathGerenciador: # Verifica se diretorio do Gerenciador de Alunos NÃO é ENCONTRADO
            messagebox.showwarning("ALERTA!", "A Máquina NÃO possui GERENCIADOR DE ALUNOS\nInicie o programa na máquina servidor!\nO Programa será encerrado...")
            Janela.destroy() #Encerra Aplicação
        else:
            datedate=datetime.datetime.now().strftime('%Y-%m-%d') #Cria uma variável com "nome" ANO-MES-DIA_HORA-MINUTO-SEGUNDO
            entradaMatricula=entryMatricula.get()
            entradaLogin=entryLogin.get()
            PathPadrão = os.path.exists('C:\COMPUWAY\ALUNOS_NOTAS')
            if not PathPadrão: # Se não encontrado o caminho C:\COMPUWAY\CPW - NOTAS - é CRIADO pasta PADRÃO
                os.makedirs('C:\COMPUWAY\ALUNOS_NOTAS')

            dadosmatricula=[]
            dadosmatricula = [int(item) for item in entradaMatricula and entradaLogin]
            if all([isinstance(item, int) for item in dadosmatricula])==True:
                if len(entradaMatricula) > 6 or len(entradaLogin) > 1:
                    entryMatricula.delete(0, 'end')
                    entryLogin.delete(0, 'end')
                    messagebox.showinfo("AVISO!", "FORMATO INCORRETO!\nMATRICULA MAIOR DO QUE 6 DIGÍTOS OU LOGIN MAIOR DO QUE 1 DÍGITO!\n \nERRO: 8000")
                else:
                    if len(entradaMatricula) == 0 and len(entradaLogin) == 0:
                        messagebox.showinfo("AVISO","INSIRA OS DADOS ANTES DE ACESSAR!\nCAMPO VAZIOS: MATRICULA e LOGIN\n \nERRO: 1000", icon='warning')
                    elif len(entradaMatricula) > 0 and len(entradaLogin) == 0:
                        messagebox.showinfo("AVISO!", "INSIRA OS DADOS ANTES DE ACESSAR!\nCAMPO VAZIOS: LOGIN\n \nERRO: 1001")
                    elif len(entradaMatricula) == 0 and len(entradaLogin) > 0:
                        messagebox.showinfo("AVISO!", "INSIRA OS DADOS ANTES DE ACESSAR!\nCAMPO VAZIOS: MATRICULA\n \nERRO: 1002")
                    else:  
                        try:
                            nomeservidor=("localhost")
                            porta=3306
                            usuario='user'
                            key='senha'
                            banco='banco'
                            
                            idmatricula=entradaMatricula+"-"+entradaLogin
                            idmatricula=str(idmatricula)
                            db_connection = mysql.connector.connect(host=nomeservidor, port=porta, user=usuario, passwd=key, database=banco)
                            cursor = db_connection.cursor()
                            try: 
                                sql = ("SELECT * from aluno WHERE matricula_aluno='"+idmatricula+"';")
                                cursor.execute(sql)
                                dadosaluno = cursor.fetchall()

                                if len(dadosaluno)==0:
                                    messagebox.showerror("ERRO!", "O aluno NÃO EXISTE!\nVerifique em seu Gerenciador de Alunos Matricula e Login Corretos!", icon='warning')
                                    entryMatricula.delete(0, 'end')
                                    entryLogin.delete(0, 'end')
                                    cursor.close()
                                    db_connection.commit()
                                    db_connection.close()

                                else:
                                    nome=(dadosaluno[0][4])
                                    sobrenome=(dadosaluno[0][5])
                                    nomealuno=(str(nome)+' '+str(sobrenome))
                                    sql = ("SELECT nome_curso, porc_teste_concluido, carga_horaria, data_teste_concluido from teste_concluido INNER JOIN aluno ON aluno.id_aluno = teste_concluido.id_aluno INNER JOIN curso ON curso.cod_curso = teste_concluido.cod_curso WHERE aula_teste_concluido='final' and matricula_aluno='"+idmatricula+"';")
                                    cursor.execute(sql)
                                    dados=[]
                                    dados=list(cursor.fetchall())

                                    if len(dados)==0:
                                        messagebox.showerror("ERRO!", "o aluno não possui cursos concluídos!\nVerifique em seu Gerenciador de Alunos Matricula e Login Corretos!", icon='warning')
                                        entryMatricula.delete(0, 'end')
                                        entryLogin.delete(0, 'end')
                                        cursor.close()
                                        db_connection.commit()
                                        db_connection.close()

                                    else:
            
                                        for i in range(len(dados)):
                                            if dados[i][1] == 100:
                                                dados[i] = dados[i][0], 10, dados[i][2] , dados[i][3]
                                            else:
                                                notaaluno=((dados[i][1])/10)
                                                dados[i] = dados[i][0], notaaluno, dados[i][2] , dados[i][3] 

                                        cabecalho=['CURSO','NOTA','CARGA (h)','DATA']
                                        dados.insert(0,cabecalho) 

                                        cursor.close()
                                        db_connection.commit()
                                        db_connection.close()

                                        pdf=FPDF(format='A4', unit='in')
                                        pdf.add_page()
                                        pdf.set_font('Times','',10.0)
                                        epw = pdf.w - 2*pdf.l_margin  
                                        col_width = epw/4
                                        th = pdf.font_size                      
                                        pdf.set_font('Helvetica','B',24.0)      
                                        pdf.set_text_color(255, 102, 0)
                                        pdf.ln(2*th)                            
                                        lh_list = [] 
                                        use_default_height = 0 #flag
                                        pdf.cell(epw, 0.0, 'COMPUWAY', align='C')            
                                        pdf.ln(3*th)
                                        pdf.cell(epw, 0.0, 'HISTORICO DE NOTAS', align='C')            
                                        pdf.ln(5*th)                         
                                        pdf.set_font('Helvetica','B',18.0)
                                        pdf.cell(epw, 0.0, 'MATRICULA.:', align='L')
                                        pdf.ln(0.0)
                                        pdf.set_text_color(0, 0, 0) 
                                        pdf.cell(epw, 0.0, idmatricula, align='C')
                                        pdf.ln(0.5)
                                        pdf.set_text_color(255, 102, 0)
                                        pdf.cell(epw, 0.0, 'NOME.:', align='L')
                                        pdf.ln(0.0)
                                        pdf.set_text_color(0, 0, 0)
                                        pdf.cell(epw, 0.0, nomealuno, align='C')
                                        pdf.ln(3*th) 
                                        pdf.set_font('Times','',10.0) 
                                        for row in dados:
                                            for datum in row: 
                                                pdf.cell(col_width, 2*th, str(datum), border=1)
                                            pdf.ln(2*th)
                                        caminhoCompuway=('C:\COMPUWAY\ALUNOS_NOTAS\Aluno_')
                                        arquivoNome=caminhoCompuway+idmatricula+"_"+datedate+".pdf"
                                        pdf.output(arquivoNome)
                                        entryMatricula.delete(0, 'end')
                                        entryLogin.delete(0, 'end')
                                        
                                        resposta = messagebox.askyesno("AVISO!", "Arquivo Notas da matricula "+idmatricula+" - "+nomealuno+ " foi criado com sucesso!\nCaminho do arquivo: "+arquivoNome+"\n\n                        Clique em SIM para visualizar.")
                                        if resposta == True:
                                            list_of_files = glob.glob('C:\COMPUWAY\ALUNOS_NOTAS\*.pdf')
                                            latest_file = max(list_of_files, key=os.path.getctime)
                                            Janela.attributes('-topmost',False)
                                            os.startfile(latest_file)
                                            
                                        else:
                                            entryMatricula.focus()

                            except mysql.connector.Error as err:
                                entryMatricula.delete(0, 'end')
                                entryLogin.delete(0, 'end')
                                messagebox.showerror("ERRO!", "NÃO FOI POSSÍVEL EXECUTAR O PROGRAMA!\n\nERRO: 9890", icon='warning')

                        except mysql.connector.Error as err:
                                entryMatricula.delete(0, 'end')
                                entryLogin.delete(0, 'end')
                                if err.errno == errorcode.ER_ACCESS_DENIED_ERROR:
                                    messagebox.showerror("ERRO!", "NÃO FOI POSSÍVEL EXECUTAR O PROGRAMA!\nSEM PERMISSÃO\nERRO: 9901", icon='warning')
                                elif err.errno == errorcode.ER_BAD_DB_ERROR:
                                    messagebox.showerror("ERRO!", "NÃO FOI POSSÍVEL EXECUTAR O PROGRAMA!\SERVIDOR NÃO EXISTE!\nERRO: 9902", icon='warning')
                                else:
                                    messagebox.showerror("ERRO!", "NÃO FOI POSSÍVEL EXECUTAR O PROGRAMA!\n\nERRO: 9900", icon='warning')
                    
    except ValueError:
        entryMatricula.delete(0, 'end')
        entryLogin.delete(0, 'end')
        messagebox.showerror("ERRO!", "NÃO FOI POSSÍVEL EXECUTAR O PROGRAMA!\nDADOS INVÁLIDOS!\nERRO: 9990", icon='warning')

Label01 = Label(Janela, padx=10, pady=10, font=titleFont, bg="#414141", fg="#f68121", text="CPW NOTA").place(x=80, y=10)
Label02 = Label(Janela, padx=2, pady=2, font=lblFont, bg="#505050", fg="#f68121", text="Insira matricula e Clique em Nota para \nimprimir Nota do aluno").place(x=50, y=62)
Label03 = Label(Janela, padx=2, pady=2, font=lblFont, bg="#505050", fg="#f68121", text="MATRICULA").place(x=40, y=120)
Label04 = Label(Janela, padx=2, pady=2, font=lblFont, bg="#505050", fg="#f68121", text="LOGIN").place(x=170, y=120)
entryMatricula = Entry(Janela, font=lblFont, width=6, bg="#646464", fg="#f68121")
entryLogin = Entry(Janela, font=lblFont, width=1, bg="#646464", fg="#f68121")
Button01 = Button(Janela,font=btnFont, padx=2, pady=2, bg="white", fg="orange", text="NOTA", command=Nota).place(x=120, y=160)
Button02 = Button(Janela, font=btnFont, padx=1, pady=1,bg="white", fg="red",text="SAIR", command=Sair).place(x=260, y=160)

entryMatricula.place(x=120, y=120)
entryLogin.place(x=220, y=120)
entryMatricula.bind('<Return>', enter_pressed)
entryLogin.bind('<Return>', enter_pressed)
entryMatricula.focus()
entryMatricula.insert(0, "")
entryLogin.insert(0, "")

Janela.mainloop()
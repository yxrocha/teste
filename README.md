# email em massa
import pyodbc, time, win32com.client, datetime as dt


teste = f'''<a href="https://wa.me/+551151161935">
                <img src="https://github.com/yxrocha/teste/blob/main/WHATSAPP_INICIO_FEIRAO.JPEG?raw=true">
            </a>
'''
contador = 0
comando_ = '''
SELECT 
A.NM_EMAIL,
B.NM_NOME,
DATEDIFF(DD,DT_VENCIMENTO,GETDATE()) ATRASO

FROM TB_CLIENTE_EMAIL A 
JOIN TB_CLIENTE B ON B.ID_CLIENTE = A.ID_CLIENTE
JOIN TB_CONTRATO D ON D.ID_CLIENTE = A.ID_CLIENTE
JOIN TB_DIVIDA E ON E.ID_CONTRATO = D.ID_CONTRATO
WHERE ID_CEDENTE = 19 
AND CAST(DT_EXPIRACAO AS DATE) > CAST(GETDATE()+2 AS DATE)
ORDER BY 3 ASC
'''
titulo = 'FEIRÃO PORTO BANK'
Dados_conexao = (
'Driver={SQL Server};'
'Server=192.168.1.3;'
'database=easycollector;'
f'UID=;'
f'pwd=;'
)
conexao = pyodbc.connect(Dados_conexao)
print('conexão bem sucedida com banco de dados')
cursor = conexao.cursor()

cursor.execute(comando_)
linhas = cursor.fetchall()
print('Quantidade de linhas: {}'.format(len(linhas)))
print('CARREGANDO....')
time.sleep(5)
for linha in linhas:
    Outlook = win32com.client.Dispatch('Outlook.application')
    mail = Outlook.CreateItem(0)
    mail.to = linha[0]
    mail.Subject = titulo
    mail.HTMLBody = f'''{teste} '''
    mail.Send()
    contador += 1
    timee = dt.datetime.now().strftime('%H:%M:%S')
    print('Email enviado para: {}'.format(linha[0]))
    print('Já foram enviados: {} e falta enviar: {}, Ultimo email foi enviado as: {}'.format(contador,len(linhas)-contador,timee))
    if timee > '21:30:00':
        time.sleep(36000)
    for c in range(1, 30):
        print('|',end='')
        time.sleep(1)

print('-=-'*13)
print('ENVIO DE EMAIL PREVENTIVO: RESULTADO')
print('-=-'*13)
print('Quantidade de enviados: {}'.format(contador))
time.sleep(9999)
print('fechando por falta de interação')
for c in range(1,5):
    print(c,end='')

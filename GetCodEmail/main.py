import imaplib
import email
import base64
import pytz
from datetime import datetime
from dateutil.parser import parse

fuso = pytz.timezone('UTC')
fusobrasilia = pytz.timezone('America/Sao_Paulo')



def getLastEmail():
    imap_server = ''
    emaill = ''
    password = ''

    imap = imaplib.IMAP4_SSL(imap_server)
    imap.login(emaill, password)
    imap.select('inbox')
    status, mensagens = imap.search(None, 'ALL')

    msg = None
    titulo = None


    if status == 'OK':
        lista_ids = mensagens[0].split()
        if lista_ids:
            ultimo_email_id = lista_ids[-1]
            status, dados = imap.fetch(ultimo_email_id, '(RFC822)')

            mensagem = email.message_from_bytes(dados[0][1])
            vect = str(mensagem['Subject']).split('?UTF-8?B?')
            dataehora = mensagem['Date']
            format = datetime.strptime(dataehora, '%a, %d %b %Y %H:%M:%S %z (%Z)')
            databrasil = format.astimezone(fusobrasilia)
            print(databrasil)

            data = databrasil.strftime('%Y-%m-%d %H:%M:%S')


            decodificados = base64.b64decode(vect[1])
            titulo = decodificados.decode('utf-8')

            for parte in mensagem.walk():
                tipo = parte.get_content_type()
                if tipo == 'text/plain':
                    corpo = parte.get_payload(decode=True).decode('ISO-8859-1')
                    msg = corpo

        else:
            print("Não há e-mails na caixa de entrada.")
    else:
        print("Erro ao buscar e-mails na caixa de entrada.")
    imap.logout()

    vector = str(msg).split('\n')
    codveficacao = None
    for i, text in enumerate(vector):
        try:
            codveficacao = int(text.strip())
            zeros = 6-len(str(abs(codveficacao)))
            codveficacao = f'{"0"*zeros}{codveficacao}'
        except:
            pass
    return titulo, codveficacao, data


atualhora = datetime.now()
title,cod, data = getLastEmail()
data = parse(data)
diferenca = atualhora-data
dhoras = diferenca.seconds // 3600
dmin = diferenca.seconds % 3600 // 60
print(dmin)





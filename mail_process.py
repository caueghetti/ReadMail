import json
import imaplib
import base64
import os
import email
from datetime import datetime
from sys import argv

def close_session():
    try:
        mail.close()
    except Exception as e:
        print(f'{datetime.now()} - CONEXAO IMAP FINALIZADA COM ERRO {e}')
        exit(1)

def mail_conect(host_,port_):
    try:
        print(f'CONECTANDO - HOST {host_} PORT {port_}')
        mail_ = imaplib.IMAP4_SSL(host=host_.strip(),port=int(port_))
        print(f'{datetime.now()} - CONEXAO IMAP ESTABELECIDA')
        return mail_
    except Exception as e:
        print(f'{datetime.now()} - ERRO AO CONECTAR AO SERVIDOR DE EMAIL : {e}')
        exit(1)

def login_mail(mail_user,mail_password):
    try:
        mail.login(mail_user,mail_password)
        print(f'LOGIN REALIZADO - {mail_user}')
    except Exception as e:
        print(f'{datetime.now()} - ERRO AO REALIZAR LOGIN NO EMAIL {mail_user} : {e}')
        exit(1)

def select_mail(mailbox,subject):
    try:
        print(f'CAIXA DE CORREIO - {mailbox}')
        print(f'ASSUNTO - {subject}')
        mail.select(mailbox)
        status, mail_id = mail.search(None, f'SUBJECT "{subject}"')
        if status.lower().strip() != 'ok':
            print('ERRO AO REALIZAR A BUSCA')
            exit(1)

        mail_id = mail_id[0].split()[-1]
        return mail_id
    except Exception as e:
        print(f'{datetime.now()} - ERRO AO SELECIONAR EMAIL {subject} NA CAIXA DE ENTRADA {mailbox} - {e}')
        exit(1)

def download_attachment(mail_id, output_path):
    try:
        status, data = mail.fetch(mail_id, "(BODY.PEEK[])")
        if status.lower().strip() != 'ok':
            print('ERRO EXTRACAO DO EMAIL')
            exit(1)
        mail_msg = email.message_from_bytes(data[0][1])
        
        for part in mail_msg.walk():
            if (part.get_content_maintype() == 'multipart' or part.get('Content-Disposition') is None):
                continue
            else:            
                fileName = part.get_filename()
                if bool(fileName):
                    print(f'ANEXO IDENTIFICADO {fileName}')
                    with open(f'{output_path}/{fileName}','wb') as arq:
                        arq.write(part.get_payload(decode=True))
                        arq.close()
                    print(f'ANEXO SALVO NO DIRETORIO {output_path}')
    except Exception as e:
        print(f'{datetime.now()} - ERRO AO REALIZAR DOWNLOAD DE ANEXO - {e}')

def read_config(id_):
    try:
        json_data = list(json.load(open('config_email.json')))
        for item in json_data:
            if item['id_'].strip().lower() == id_.strip().lower():
                return item
            else:
                continue
        print(f'{datetime.now()} - CONFIGURACOES NAO ENCONTRADAS NO ARQUIVO DE PARAMETROS - {id_}')
        exit(0)
    except Exception as e:
        print(f'{datetime.now()} - ERRO AO LER ARQUIVO DE CONFIGURACOES (SESSION - {id_}): {e}')
        exit(1)

if __name__ == "__main__":
    info_session = read_config(str(argv[1]))
    mail = mail_conect(info_session['imap_url'],info_session['imap_port']) #conecta no servidor de e-mail
    user_=str(info_session['user']).strip()
    password_=str(info_session['password']).strip()
    print(user_)
    print(password_)
    login_mail(user_,password_) #realiza login no email
    id__mail = select_mail(info_session['mailbox'],info_session['subject'])
    download_attachment(id__mail,info_session['dir_path']) #realiza leitura do email e download de anexo caso haja
    close_session() # finaliza a sessao



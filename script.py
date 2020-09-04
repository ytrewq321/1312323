import smtplib                                              
import os

import mimetypes                                            
from email import encoders                                  
from email.mime.base import MIMEBase                        
from email.mime.text import MIMEText                        
from email.mime.image import MIMEImage                      
from email.mime.audio import MIMEAudio                      
from email.mime.multipart import MIMEMultipart              


def send_email(addr_to, msg_subj, msg_text, files):
    addr_from = "vladislavkotkin565358@gmail.com"                         
    password  = "5653485649Vladik"                                  

    msg = MIMEMultipart()                                  
    msg['From']    = addr_from                             
    msg['To']      = addr_to                              
    msg['Subject'] = msg_subj                               

    body = msg_text                                        
    msg.attach(MIMEText(body, 'plain'))                     

    process_attachement(msg, files)

    server = smtplib.SMTP_SSL('smtp.gmail.com', 465)        
    #server.starttls()                                      
    #server.set_debuglevel(True)                            
    server.login('vladislavkotkin565358@gmail.com', '5653485649Vladik')                     
    server.send_message(msg)                                
    server.quit()                                           
    
def process_attachement(msg, files):                        
    for f in files:
        if os.path.isfile(f):                               
            attach_file(msg,f)                              
        elif os.path.exists(f):                             
            dir = os.listdir(f)                             
            for file in dir:                                
                attach_file(msg,f+"/"+file)                 

def attach_file(msg, filepath):                            
    filename = os.path.basename(filepath)                  
    ctype, encoding = mimetypes.guess_type(filepath)        
    if ctype is None or encoding is not None:               
        ctype = 'application/octet-stream'                  
    maintype, subtype = ctype.split('/', 1)                 
    if maintype == 'text':                                  
        with open(filepath) as fp:                          
            file = MIMEText(fp.read(), _subtype=subtype)    
            fp.close()                                      
    elif maintype == 'image':                               
        with open(filepath, 'rb') as fp:
            file = MIMEImage(fp.read(), _subtype=subtype)
            fp.close()
    elif maintype == 'audio':                               
        with open(filepath, 'rb') as fp:
            file = MIMEAudio(fp.read(), _subtype=subtype)
            fp.close()
    else:                                                   
        with open(filepath, 'rb') as fp:
            file = MIMEBase(maintype, subtype)              
            file.set_payload(fp.read())                     
            fp.close()
            encoders.encode_base64(file)                   
    file.add_header('Content-Disposition', 'attachment', filename=filename) 
    msg.attach(file)                                        
    
addr_to   = "kotkin.vladislav@mail.ru"                                
files = ["test.xlsx"]
send_email(addr_to, "Тема сообщения", "Текст сообщения", files)

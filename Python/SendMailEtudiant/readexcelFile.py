from pandas_ods_reader import read_ods
import smtplib
import sys
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText


def sendEmail(message, studentEmail):
    
    receivers = studentEmail
    sender =  "youremail"
    # Create message container - the correct MIME type is multipart/alternative.
    msg = MIMEMultipart('alternative')
    msg['Subject'] = "yourSubject"
    msg['From'] = sender
    msg['To'] = studentEmail
    part1 = MIMEText(message, 'plain')
    msg.attach(part1)
    
    password ='yourpassword'
    server = smtplib.SMTP("smtp.serverDomain.fr", 587) #connect to SMTP object that has methods to send email
    server.ehlo() #establish a connection with the server by sending a hello message
    server.starttls() #start a TLS encryption service(uses port 587)
    server.login(sender, password) #login to your mail service
    server.sendmail(sender, studentEmail, msg.as_string()) #send email takes 3 args 
    server.quit()


def readNotes ():
    
    path = "yourPath.ods"
    # load a sheet based on its index (1 based)
    sheet_name = "Feuille2"
    df = read_ods(path, sheet_name)

    for index, cell in df.iterrows():
    #    print(cell[0], cell[1],cell[2])
        message = (f"Bonjour,"+
                   f"\n\nComme promis, voici votre note : {cell[2]}"  
                   f"\nLes remarques : "
                   f"\n     Explication Note Ex1 : {cell[4]} "
                   f"\n     Explication Note Ex2 : {cell[6]} "               
                   f"\n\nCordialement")
        
        print(message)
        sendEmail(message, cell[0])
        sendEmail(message, cell[1]) 
        break

print("Start Read file Notes")
readNotes()
print("All successful done !!")


#NioulBoy
import smtplib

#conn =smtplib.SMTP('smtp.gmail.com',587)
conn =smtplib.SMTP('smtp-mail.outlook.com',587)
type(conn)
conn.ehlo()
conn.starttls()

print("Entrer votre login svp:")
log = input()
print("Entrer votre mdp svp:")
mdp = input()
conn.login(log,mdp)

receiver = "mbayedione10@gmail.com"
message = """\
Subject: Requete SQL Mensuel

This message is sent from Python.Success"""

conn.sendmail('mbayedione10@hotmail.fr',receiver,message,)
conn.quit()



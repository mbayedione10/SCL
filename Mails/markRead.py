#NioulBoy
#https://stackoverflow.com/questions/51596061/find-the-subject-of-a-specific-email-imap-python
#https://desktop.arcgis.com/fr/arcmap/latest/analyze/executing-tools/scheduling-a-python-script-to-run-at-prescribed-times.htm
import  imaplib

import getpass

#connexion

mail = imaplib.IMAP4_SSL('Outlook.office365.com')

print("Entrer votre login svp:")
log = input()
print("Entrer votre mdp svp:")
mdp = input()

mail.login(log,mdp)
mail.select()
#filtrer les mails
typ, data = mail.search(None,'FROM', "sender@mail.com", "UNSEEN", 'SUBJECT', "Success")
#type, data =mail.delete(None,'FROM', "sender@mail.com", "UNSEEN", 'SUBJECT', "Success")

#Lire
for num in data[0].split():
    typ, data = mail.fetch(num, '(RFC822)')
    mail.store(num, '+FLAGS', '\\Deleted')
    print("deleted succesfully")


mail.close()  #close the mail box
mail.logout() #logout

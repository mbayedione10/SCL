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

#search specific mails
typ, data = mail.search(None,'FROM', "f.sow@sesam.sn", "UNSEEN", 'SUBJECT', "Success")
#typ, data = mail.search(None,'FROM', "f.sow@sesam.sn", "SEEN", 'SUBJECT', "Success")


#Lire
for num in data[0].split():
    #Read
    typ, data = mail.fetch(num, '(RFC822)')
    #Delete
    #mail.store(num, '+FLAGS', '(\\Deleted)')

#print("deleted succesfully")


mail.close()  #close the mail box
mail.logout() #logout

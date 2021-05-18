#NioulBoy
#https://stackoverflow.com/questions/51596061/find-the-subject-of-a-specific-email-imap-python
#https://desktop.arcgis.com/fr/arcmap/latest/analyze/executing-tools/scheduling-a-python-script-to-run-at-prescribed-times.htm
import  imaplib
import datetime
import getpass

date = (datetime.date.today() - datetime.timedelta(days=15)).strftime("%d-%b-%Y")
mail = imaplib.IMAP4_SSL('Outlook.office365.com')

log = input("Entrer votre login svp: ")
mdp = getpass.getpass()

mail.login(log,mdp)
mail.select()

#Search mailbox for matching messages.
#'mark' and 'delete' are space separated list of matching message numbers.
typ1, mark = mail.search(None,'FROM', "mailfrom@mail.com", "UNSEEN", 'SUBJECT', "Success")
typ2, delete = mail.search(None,'FROM', "mailfrom@mail.com", "SEEN", 'SUBJECT', "Success","SENTBEFORE", date)


for num in mark[0].split():py
    typ1, mark = mail.fetch(num, '(RFC822)')

for sup in delete[0].split():
    mail.store(sup, '+FLAGS', '(\\Deleted)')
print("Mark and delete done")

mail.close()  #close the mail box
mail.logout() #logout

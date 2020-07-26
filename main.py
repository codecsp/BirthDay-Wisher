import pandas as pd
import datetime
import smtplib
from email.message import EmailMessage

GmailID = ''
gmailPass = ''

def sendEmail(to,sub,mesg):
	try:
		msg = EmailMessage()
		msg['Subject'] = sub
		msg['From'] = GmailID
		msg['To'] = to
		msg.set_content(str(mesg))

		s=smtplib.SMTP('smtp.gmail.com',587)
		s.starttls()
		s.login(GmailID,gmailPass)
		# s.sendmail(GmailID,to,message)
		# print("msg sent to :"+ to)
		s.send_message(msg)
		s.quit()

	except SMTPException:
		print ("Error: unable to send email")

   

if __name__ == '__main__':
	df = pd.read_excel("data.xlsx")
	#print(df)
	today = datetime.datetime.now().strftime("%d-%m") #day and month only
	yearnow = datetime.datetime.now().strftime("%Y")
	
	#print(type(today))
	writeind = []

	for index, item in df.iterrows():
		bday = item['Birthday'].strftime("%d-%m")
		to= item['Email']
		msg = item['Dialogue']
		sub= "HappY Birthday "+ item['Name']
		#print(index, item['Birthday'])
		#print(bday)
		if(today==bday) and yearnow not in str(item['Year']):
			sendEmail(to,sub,msg)
			writeind.append(index)


	#print(writeind)
	for i in writeind:
		yr = str(df.loc[i,'Year'])
		df.loc[i,'Year'] = yr+","+yearnow;
		#print(df.loc[i,'Year'])

	#print(df)
	df.to_excel('data.xlsx',index=False)




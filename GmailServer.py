import xlrd
import xlwt,getpass
import time
import smtplib
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
from email.mime.base import MIMEBase
from email import encoders
import random
from datetime import datetime
import xlutils
from xlutils.copy import copy
count_email=0
try:
	book1=xlrd.open_workbook("gmaildata.xls")
	ws5=book1.sheet_by_index(0)
	c1=ws5.cell(-1,5)
	i=int(c1.value)
	i=i+1
except (IndexError,FileNotFoundError):
	i=0
wb=xlwt.Workbook()
ws1=wb.add_sheet("sheet1")
check=0
print("Welcome to gmail server !!")
while True:
	time.sleep(1)
	print(" Enter 1 for Sign up :","\n","Enter 2 for Sign in : ")
	tick=int(input())
	if tick==1:
		name=input("Enter Your Full Name :")
		dob=input("Enter your Date of Birth in dd/mm/yyyy format :")
		while True:
			phone=input("Enter your 10 digit Mobile no :")
			if len(phone)!=10:
				print("Your Mobile no is not 10 digit Please Enter correct : ") 
				continue
			else:
				break
		while True:
			mail=input("Enter a new mail id  in the format @gmail.com :")
			if mail[-10:]!="@gmail.com":
				print("Your mail id is not in format of @gmail.com Please Enter in  valid format :")
				continue
			else:
				break
		while True:
			pass0=getpass.getpass("Enter a Password :")
			pass1=getpass.getpass("Re Enter your Password for confirmation :")
			if pass0!=pass1:
				print("Your Enter Password is not match Please Enter Valid Password :")
				continue
			else:
				break
		
		if i==0:
			wb=xlwt.Workbook()
			ws1=wb.add_sheet("sheet1")
			ws1.write(i,0,name)
			ws1.write(i,1,dob)
			ws1.write(i,2,phone)
			ws1.write(i,3,mail)
			ws1.write(i,4,pass0)
			ws1.write(i,5,i)
			wb.save("gmaildata.xls")
		else :
			book=xlrd.open_workbook("gmaildata.xls")
			wb=copy(book)
			wb.get_sheet(0).write(i,0,name)
			wb.get_sheet(0).write(i,1,dob)
			wb.get_sheet(0).write(i,2,phone)
			wb.get_sheet(0).write(i,3,mail)
			wb.get_sheet(0).write(i,4,pass0)
			wb.get_sheet(0).write(i,5,i)
			wb.save("gmaildata.xls")
		i=i+1
		print("Thank You !! You are Successfully Sign up !!")
		continue
	count_email=i-1
	if tick==2:
		while True:
			user_name=input("Enter Your gmail id in format of @gmail.com : ")
			wb=xlrd.open_workbook("gmaildata.xls")
			ws2=wb.sheet_by_index(0)
			flag=0
			for i in range(0,count_email+1):
				c2=ws2.cell(i,3)
				if c2.value==user_name:
					flag=1
					break
			if flag==0:
				print("Gmail is invalid Please enter valid email id--")
			else:
				break
		countie=0
		while True:
			password=getpass.getpass("Enter your Password : ")
			wb=xlrd.open_workbook("gmaildata.xls")
			ws2=wb.sheet_by_index(0)
			c3=ws2.cell(i,4)
			flag=0
			if c3.value==password:
				break
			else:
				print("Password is invalid Please enter valid Password--")
				countie+=1
				if countie==3:
					print("You try more chances !!","\n","Enter 1 for Forget Password ","\n","Enter 2 for  Re Enter your Password after 1 minute :")
					fp=int(input())
					if fp==1:
						while True:
							phn=input("Enter Mobile no that You entered at the time of Sign up : ")
							wb=xlrd.open_workbook("gmaildata.xls")
							ws2=wb.sheet_by_index(0)
							c4=ws2.cell(i,2)
							if c4.value==phn:
								while True:
									new_pass=getpass.getpass("Enter New Password : ")
									confirm=getpass.getpass("Re Enter Password for Confirmation : ")
									if new_pass==confirm:
										book=xlrd.open_workbook("gmaildata.xls")
										wb=copy(book)
										wb.get_sheet(0).write(i,4,new_pass)
										wb.save("gmaildata.xls")
										print("Thank you Your Password is changed Successfully!! Please Login With New Password-- ")
										break
									else:
										print("Password Enter is not match !! Please Re Enter Your Password--")
										continue
								break
							else:
								print("Mobile no is invalid !!  Please Enter a Valid Mobile no -- ")
								continue
					else :
						print("Please wait for 1 minutes.... !! ")
						time.sleep(60)
						countie=0
				else:
					continue
# the email id in the excel sheet are not valid gmail id ( for security purpose) so I have not used them in the upcoming program.					
		another_mail=input("Enter another gmail id for sending OTP : ")
		otp=random.randint(100000,1000000)
		server= smtplib.SMTP("smtp.gmail.com",587)
		#print("Establishing the connection")
		server.starttls()
		#print("Login in ")
		main_email="ankurgupta9352@gmail.com"
		server.login(main_email,"password") # for security purpose I have not written my password
		#print("sending mail ")
		message="This is your verification otp-- "
		msg=message+str(otp)
		server.sendmail("ankurgupta9352@gmail.com",another_mail,msg)
		server.quit()
		cou=0
		while True:
			otp1=int(getpass.getpass("Enter OTP : "))
			if otp1==otp:
				print("You are Successfully login in your Gmail !! ")
				time.sleep(1)
				break
			else:
				print("Oops !! OTP is invalid Enter a valid OTP -")
				cou+=1
				print(3-cou,"Chance is left!!")
				if cou==3:
					print("Sorry !! Your Gmail id is blocked PLease Contact to Customer Care No = +91 1234567890"  )
					book=xlrd.open_workbook("gmaildata.xls")
					wb=copy(book)
					wb.get_sheet(0).write(i,3,"blocked")
					wb.save("gmaildata.xls")
					exit()
				else:
					continue
		while True:
			print(" Enter 1 for sending a Mail :","\n","Enter 2 for logOut :")
			send_mail=int(input())
			if send_mail==1:
				op=input("Enter a mail id where you want to send a mail : ")
				msg = MIMEMultipart()
				msg['From'] = main_email
				msg['To'] = op
				sub=input("Write a subject of your mail")
				msg['Subject'] = sub
				body = input("Body_of_the_mail")
				msg.attach(MIMEText(body, 'plain'))
				ques=int(input("Press 1 to attach file to your mail else Press 2"))
				if ques==1:
					filename =input("Enter the file name")
					location=input("Enter the file location")
					attachment = open(location, "rb")
					p = MIMEBase('application', 'octet-stream')
					p.set_payload((attachment).read())
					encoders.encode_base64(p)
					p.add_header('Content-Disposition', "attachment; filename= %s" % filename)
					msg.attach(p)
				server= smtplib.SMTP("smtp.gmail.com",587)
				#print("Establishing the connection")
				server.starttls()
				#print("Login in ")
				server.login(main_email,"password") # for security purpose I have not written my password
				text = msg.as_string()
				print("sending mail .. ")
				server.sendmail("ankurgupta@gmail.com",op,text)
				server.quit()
				print("Mail is sent !! ")
				continue
			else:
				print("Thank you !!")
				exit()
					
	else:
		print("You Enter a Invalid Option Please Enter Valid Option--")
		continue

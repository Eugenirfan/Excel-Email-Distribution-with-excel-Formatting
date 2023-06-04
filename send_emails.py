#libraries
import pandas as pd
import glob
# getting excel files from Directory Desktop
path = r"S:t\Irfan\python projects\XSNM by Name"

# read all the files with extension .xlsx i.e. excel 
filenames = glob.glob(path + "\*.xlsx")

#print('File names:', filenames)
file_name=[]
for file in filenames:
    char1= 'Name\\'
    char2= '.xlsx'

    #extract name from filenames
    name=file[file.find(char1)+5: file.find(char2)]
    #create a list of all supplier names called file_name
    file_name.append(name)
    
    
print(file_name)


    #adding column names to file and supplier name
f={'File': filenames, 'Name':file_name}
df=pd.DataFrame(f)
pd.set_option("display.max_colwidth", -1)
df.head()

#create dummy emails
list_email=[]
mail='email'
com= '@email.com'
for i in range(len(filenames)):
    emails=mail+str(i)+com
    list_email.append(emails)
#print(list_email)  

#libraries
import os
import win32com.client
subject = "excess"
outlook = win32com.client.Dispatch('outlook.application')
for index, row in df.iterrows():
    mail_send = outlook.CreateItem(0)
    mail_send.To = row['email']
    attachment  = row['File']
    mail_send.Attachments.Add(attachment)
    mail_send.Subject = 'Excess No Move' +" "+ row['Name']
    #mail_send.Body = 'Dear \n\n'\
                    #'.\n\n'\
                    #'Thank you\n\n'\
                    #'Irfan\n\n'
    
    #mail_send.CC='mohamedirfan.suffeerahmed@gmail.com'
    mail_send.Send()
    

    

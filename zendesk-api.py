import requests
import boto3
from openpyxl import Workbook
from openpyxl.styles import Font
from email.mime.text import MIMEText
from email.mime.application import MIMEApplication
from email.mime.multipart import MIMEMultipart


def get_details():
 client = boto3.client('ses',region_name='us-east-2',aws_access_key_id='xxxxxx',aws_secret_access_key='xxxxxx')
 #boto3 is used for sending amazon ses emails based on key and secret
 
 for a in range(1):
  #query to get change request form for the last one week.
  url = 'https://example.zendesk.com/api/v2/search.json?query=created>1week form:"Change Request"&page='+str(a)
 
  # authenticate with zendesk api token
  response = requests.get(url, auth=('user@domain.com/token','xxxxxx'))
 
  response = response.json()

  #excel sheet workbook
  work = Workbook()
  sheet = work.active
  

  #initializing columns with font
  sheet['A1']="Domain"
  sheet['B1']="Subject"
  sheet['C1']="Requester Name"
  sheet['D1']="Requester Email"
  sheet['E1']="Project Manager"
  sheet['F1']="Jira Tickets "
  sheet['A1'].font = Font(bold=True,size=16,color='1632e3')
  sheet['B1'].font = Font(bold=True,size=16,color='1632e3')
  sheet['C1'].font = Font(bold=True,size=16,color='1632e3')
  sheet['D1'].font = Font(bold=True,size=16,color='1632e3')
  sheet['E1'].font = Font(bold=True,size=16,color='1632e3')
  sheet['F1'].font = Font(bold=True,size=16,color='1632e3')

  row_count=0
 
  #extracting individual tickets 

  for urls in response['results']: 
     response1 = requests.get(urls['url'], auth=('user@domain.com/token','xxxxxx'))
     response1 = response1.json()
     row_count=row_count+1  #increase rowcount after fetching each url's
        #getting subject name of ticket  type=str

     print("----------------------------------------------------------------------------------------")   
     print("\n Subject is : ")   
     ticket_subject = response1['ticket'].get("subject")
     print(ticket_subject)

     #getting requester name and email type=str

     print("\n Requester name is : ")
     requester_id = response1['ticket'].get("requester_id")
     requester_url = 'https://example.zendesk.com/api/v2/users/'+str(requester_id)
     response2 = requests.get(requester_url, auth=('user@domain.com/token','xxxxxx'))
     response2 =response2.json()
     ticket_requester_name = response2['user'].get("name")
     print(ticket_requester_name)
     print("\n Requester email id is : ")
     ticket_requester_email=response2['user'].get("email")
     print(ticket_requester_email)

        #getting manager name  type=str

     dict1 = next( item for item in response1['ticket']['custom_fields']  if item["id"] == int("360010131412"))
     print("\n Project Manager is : ")
     ticket_project_manager = dict1.get("value")
     print(ticket_project_manager)

     #getting domain name  this is list 

     dict1 = next( item for item in response1['ticket']['custom_fields']  if item["id"] == int("360010241911"))
     print("\n Domain  : ")
     ticket_domain_name = dict1.get("value")
     print(ticket_domain_name)
     

     #getting jira tickets type=str

     print("\n Jira tickets are : \n ")
     dict1 = next( item for item in response1['ticket']['custom_fields']  if item["id"] == int("360010241951"))
     ticket_jira = dict1.get("value")
     print(ticket_jira)
    
     
      # populating cell values in excell sheet with column 
     for colns in range(6):
        celldata = sheet.cell(row=row_count+1, column=colns+1)
        if colns == 0 :
        	celldata.value = str(ticket_domain_name)
        elif colns == 1 :
            celldata.value = ticket_subject
        elif colns == 2 :
            celldata.value = ticket_requester_name
        elif colns == 3 :
            celldata.value = ticket_requester_email
        elif colns == 4 :
            celldata.value = ticket_project_manager
        else :
            celldata.value = ticket_jira

            
 # after data is populated saving into excel file

 work.save(filename="zendesk-data.xlsx")   
 
 #message subject 
 message = MIMEMultipart()
 message['Subject'] = 'Email subject'
 message['From'] = 'sender@domain.com'
 message['To'] = ', '.join(['receipient@domain.com', 'receipient1@domain.com'])
 

 #message body
 part = MIMEText('Hello Team , \n Zendesk tickets details for this week are attached. \n Thanks  ', 'html')
 message.attach(part)

   # File name provided
 part = MIMEApplication(open('zendesk-data.xlsx', 'rb').read())
 part.add_header('Content-Disposition', 'attachment', filename='zendesk-data.xlsx')
 message.attach(part)

 email_sending = client.send_raw_email(
    Source=message['From'],
    Destinations=['receipient@domain.com', 'receipient1@domain.com'],
    RawMessage={ 'Data': message.as_string() }
  )
     


if __name__ == "__main__" :
	get_details()









 	
 	



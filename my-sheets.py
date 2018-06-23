import gspread
#import datetime
from oauth2client.service_account import ServiceAccountCredentials
#from datetime import timedelta

# Download the helper library from https://www.twilio.com/docs/python/install
import twilio
import twilio.rest
from twilio.rest import Client

#Datetime import
from datetime import date
from datetime import time
from datetime import datetime
from datetime import timedelta

# use creds to create a client to interact with the Google Drive API
#scope = ['https://spreadsheets.google.com/feeds']
scope = ['https://spreadsheets.google.com/feeds',
         'https://www.googleapis.com/auth/drive']

# creds on windows in office
creds = ServiceAccountCredentials.from_json_keyfile_name('c:/Users/dimit/.thonny/client_secret.json', scope)

# creds on laptop anche
#creds = ServiceAccountCredentials.from_json_keyfile_name('/Volumes/Macintosh HD/Users/dimitri/CloudStation/001-workflofApp/client_secret.json', scope)

gs = gspread.authorize(creds)

# Find a workbook by name and open the first sheet
# Make sure you use the right name here.
sheet = gs.open("orteliusWorkflow").sheet1
 
# Extract and print all of the values
all_bookings = sheet.get_all_records()

#count number of bookings in sheet (#rows)
number_of_bookings = (len(all_bookings))

#loop to go trough the booking
#i=0
#while i < number_of_bookings:
#    print (all_bookings[i]['Name'],all_bookings[i]['End'])
#    i=i+1

print ("total # of bookings =", (number_of_bookings))

# >>> --- START time checking and message sedning routine 
i=0
while i < number_of_bookings:
    # set-up date calculation variables
    date = datetime.strptime((all_bookings[i]['Start']),'%d %B %Y')
    today = datetime.today()
    Daysleft = int(str((date-today) + timedelta(days=1)).split(' ')[0])
    
    # print statement only used for control output in code - uncomment when editing - >>> print ("daysleft: " + str(Daysleft))
    
    #check if arrivaldate in booking is three days from now, in which case the IF loop starts
    if (Daysleft) <= 3 and (Daysleft) > 0 and (all_bookings[i]['confirmed'])=='yes' and (all_bookings[i]['send_reminder'])=='yes':
        
        #load booking information in strings
        vor = (all_bookings[i]['Vorname'])
        name = (all_bookings[i]['Name'])
        start = (all_bookings[i]['Start'])
        end = (all_bookings[i]['End'])
        night = (all_bookings[i]['nights'])
        tel = (all_bookings[i]['telefon'])
        pers = (all_bookings[i]['pers'])
        comments = (all_bookings[i]['comment'])
        
        
        
        #send message to guest
        print (("Hello ") + (vor) + (" in ") + str(Daysleft) + (" days you will be visiting our appartment and stay from ") + str(start) + (" to ") + (end) + (" please contact Marjolein on +31 11addnumber11 to let her know when you will be arriving"))
    
        #send message to Marjolein
        print (("Hello Marjolein in ") + str(Daysleft) + (" days ") + (vor) + (" ") + (name) + (" will be visitng our appartment and stay from ") + (start) + (" to ") + (end) + (" ") + (vor) + (" will come with ") + str(pers) + (" person(s) in total and can be reached under tel: ") + (tel) + (" further relevant information:  ") + (comments))
    
        #>>> Start Twilio routine
            # Your Account Sid and Auth Token from twilio.com/console
        account_sid = 'ACb7315404d8e4df9a8104e35f630478dc'
        auth_token = 'bfc6dc7f8d448e4cf11f2ad2d86ff90a'
        client = Client(account_sid, auth_token)
        
        message = client.messages.create(
                                      body=(("Hello ") + (vor) + (" in ") + str(Daysleft) + (" days you will be visiting our appartment and stay from ") + str(start) + (" to ") + (end) + (" please contact Marjolein on +31 11addnumber11 to let her know when you will be arriving")),
                                      from_='+16476994815',
                                      to='+4917657168394'
                                  )
        print(message.sid)
        #<<< End Twilio routine
        
        #Set "reminder send" flag to "done" for the relevant booking (based on var (i)) 
        sheet.update_acell(('s'+str(i+2)), 'done')
    i=i+1
# >>> --- END time checking and message sedning routine









    
    
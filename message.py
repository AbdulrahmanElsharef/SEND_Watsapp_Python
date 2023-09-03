import pywhatkit
import datetime
import openpyxl
# currentime = datetime.datetime.now()
# hours = currentime.strftime('%H')
# minutes = currentime.strftime('%M')

# phone="+201111831251"
# message="عاوزين نفيش الهوامش"
# pywhatkit.sendwhatmsg(phone, message, 0, 9)

# # # Load the Excel sheet
# excel_file = 'F:\Al_Tawheed/contact.xlsx'
# sheet_name = 'Sheet1'
# df = pd.read_excel(excel_file, sheet_name=sheet_name)


# currentime = datetime.datetime.now()
# hours = currentime.strftime('%H')
# minutes = currentime.strftime('%M')
# second = currentime.strftime('%S')
# # print(f'time hour is {int(hours)}')
# # print(f'time minutes is {int(minutes)}')
# # print(f'time second is {int(second)}')
# pywhatkit.sendwhatmsg('+201153156090', 'بسم الله ماشاء الله  ', 1, 57, 15, True, 2)



# # Specify the path to your Excel file
excel_file = 'F:\Al_Tawheed\watsapp/contact.xlsx'

# # Load the Excel workbook
workbook = openpyxl.load_workbook(excel_file)

# # Select the sheet you want to read
sheet = workbook['whatsapp']  # Replace 'Sheet1' with your desired sheet name
v_range=sheet['C2':'D5']

for x,y in v_range:
    currentime = datetime.datetime.now()
    hours = currentime.strftime('%H')
    minutes = currentime.strftime('%M')
    seconds = currentime.strftime('%S')
    phone=x.value
    msg=y.value
    pywhatkit.sendwhatmsg(phone, msg, int(hours), int(minutes)+1,15,True)
    # minutes+=1
    print(f'send message to {phone} ')

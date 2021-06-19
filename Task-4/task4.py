import ClointFusion as cf
cf.OFF_semi_automatic_mode()
import time
excelfile = r"C:\Users\chirag mali\Desktop\clointfusion task\task 4\Email.xlsx"

def getData(excelPath='',sheet_name='',columnName=''):
    data = []
    row, column =cf.excel_get_row_column_count(excel_path=excelPath,sheet_name=sheet_name)
    for i in range(row-1):
        temp = cf.excel_get_single_cell(excel_path=excelPath,sheet_name=sheet_name,columnName=columnName,cellNumber=i)
        if str(temp) != 'nan':
            data.append(temp)
    return data

row, column = cf.excel_get_row_column_count(excel_path= excelfile, sheet_name='Sheet1')
headers = cf.excel_get_all_header_columns(excel_path= excelfile, sheet_name='Sheet1')
mailId = getData(excelPath= excelfile, sheet_name='Sheet1', columnName=headers[0])
names = getData(excelPath= excelfile, sheet_name='Sheet1', columnName=headers[1])
hackathon_version = getData(excelPath= excelfile, sheet_name='Sheet1', columnName=headers[2])
date = getData(excelPath= excelfile, sheet_name='Sheet1', columnName=headers[3])

def set_Version_date(msg,hackathon_version,date):
    cf.message_counter_down_timer('Set version & date from the text file',start_value=3)
    hackword = 'Hackathon'
    dateword = 'Time:'
    for i in range(len(msg)):
        line = msg[i]
        lineList = line.split()
        if hackword in lineList:
            index = lineList.index(hackword) + 1
            lineList[index] = str(hackathon_version[0])
            msg[i] = ' '.join(lineList)
            
        if dateword in lineList:
            index = lineList.index(dateword)
            for k in range(3):
                lineList.pop(index + 1)
            lineList[index+1] = date[0]
            msg[i] = ' '.join(lineList)
    return msg

def create_msg(msg,mailId,name):
    cf.message_counter_down_timer(f'Create message for {name}',start_value=2)
    lineList = msg[0].split()
    lineList[1] = f'{name},'
    msg[0] = ' '.join(lineList)
    print(msg)
    
    with open(f'{name}.txt','w') as f:
        f.writelines(msg)
    f.close()

def setupMail():
    cf.launch_website_h('https://outlook.live.com/owa/')
    cf.message_counter_down_timer('Wait Outlook is opening',start_value=3)
    cf.browser_wait_until_h(text='Sign in')
    cf.browser_mouse_click_h(User_Visible_Text_Element='Sign in')
    
#     login details with gui functions
    username = cf.gui_get_any_input_from_user("Enter your Mail id")
    cf.key_write_enter(username)

    password = cf.gui_get_any_input_from_user("Enter your Password", password=True)
    cf.key_write_enter(password)
        
    return True

def sendMail(name,mailId):
    cf.message_counter_down_timer('Now Set Up mail',start_value=3)
    cf.browser_wait_until_h(text='New Message')
    cf.browser_mouse_click_h(User_Visible_Text_Element='New Message')
    
    cf.browser_wait_until_h(text='To')
    cf.browser_write_h(Value=mailId,User_Visible_Text_Element='To')
    
    cf.browser_write_h(Value='Hackathon Invitation mail TEST',User_Visible_Text_Element='Add a Subject')
    
    cf.mouse_click(x=1196,y=684,single_double_triple='single')
    with open(f'C:/Users/chirag mali/Desktop/clointfusion task/task 4/{name}.txt','r') as f:
        for line in f.readlines():
            cf.key_write_enter(strMsg=line,delay=0.1,key=' ')
    f.close()

    cf.browser_mouse_click_h(User_Visible_Text_Element='Send')

if __name__ == "__main__":
    
    with open(file='Email.txt',mode='r') as f:
        msg = f.readlines()
    f.close()

    newMsg = set_Version_date(msg,hackathon_version,date)
    for i in range(row-1):
        create_msg(newMsg,mailId[i],names[i])

    if setupMail():
        for i in range(row-1):
            sendMail(names[i],mailId[i])

    cf.message_counter_down_timer("Closing Browser",start_value=10)
    cf.browser_quit_h()
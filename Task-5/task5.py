import ClointFusion as cf
import os
import time

from numpy import string_

cf.OFF_semi_automatic_mode()
cf.window_show_desktop()


cf.launch_website_h("https://avinashtechlvr.github.io/ClointFusion-Training-Task-6/")
cf.scrape_save_contents_to_notepad('C:/Users/chirag mali/Desktop/clointfusion task/task 5')
cf.browser_quit_h()


cf.launch_any_exe_bat_application('notepad-contents.txt')
cf.key_press('ctrl+a')
cf.key_press('ctrl+c')
cf.window_close_windows('notepad')
cf.excel_create_file("C:/Users/chirag mali/Desktop/clointfusion task/task 5","noteToexcel")

cf.launch_any_exe_bat_application('noteToexcel.xlsx')
cf.key_press('ctrl+v')
cf.key_press('ctrl+s')
cf.key_press('alt+f3')
cf.key_write_enter("A1")
for i in range(2):
    cf.key_press("ctrl+-")
    cf.key_press('Down')
    cf.key_press('Down')
    cf.key_hit_enter()
for i in range(5):
    cf.key_press('Down')
    cf.key_press("ctrl+-")
    cf.key_press('Down')
    cf.key_press('Down')
    cf.key_hit_enter()

   
cf.key_press("ctrl+s")
cf.window_close_windows('Excel')
cf.launch_website_h("https://www.xe.com/currencyconverter/convert/?Amount=2000&From=GBP&To=CAD")
for j in range(5):

    From_currency = cf.excel_get_single_cell('C:/Users/chirag mali/Desktop/clointfusion task/task 5/noteToexcel.xlsx',columnName="From",cellNumber=j)
    To_currency=cf.excel_get_single_cell('C:/Users/chirag mali/Desktop/clointfusion task/task 5/noteToexcel.xlsx',columnName="To",cellNumber=j)
    Amount=str(cf.excel_get_single_cell('C:/Users/chirag mali/Desktop/clointfusion task/task 5/noteToexcel.xlsx',columnName="Amount",cellNumber=j))
    # print(From_currency)
    # print(To_currency)
    # print(Amount)
    time.sleep(3)
    cf.mouse_click(350,827,single_double_triple='single')
    cf.browser_mouse_click_h('From')
    cf.key_write_enter(From_currency)
    cf.key_press('Tab')
    cf.key_press('Tab')
    cf.key_write_enter(To_currency)
    cf.browser_mouse_click_h("Amount")
    cf.key_write_enter(Amount)
    cf.key_press('ctrl+a')
    cf.key_write_enter(Amount)
    Amount=cf.browser_locate_element_h('//*[@id="__next"]/div[2]/div[2]/section/div[2]/div/main/form/div[2]/div[1]/p[2]',get_text=True).split()

    Amount[0]=Amount[0].replace(',','')
    print(Amount[0])
    cf.excel_set_single_cell('C:/Users/chirag mali/Desktop/clointfusion task/task 5/noteToexcel.xlsx',columnName="Converted",cellNumber=j,setText=Amount[0])
cf.browser_quit_h()

#Signin outlook
cf.launch_website_h('https://outlook.live.com/owa/')
cf.message_counter_down_timer('Wait Outlook is opening',start_value=3)
cf.browser_wait_until_h(text='Sign in')
cf.browser_mouse_click_h(User_Visible_Text_Element='Sign in')
    
#     login details with gui functions
username = cf.gui_get_any_input_from_user("Enter your Mail id")
cf.key_write_enter(username)

password = cf.gui_get_any_input_from_user("Enter your Password", password=True)
cf.key_write_enter(password)
time.sleep(5)
cf.browser_mouse_click_h('New Message')
time.sleep(3)
cf.key_write_enter('chiragm129@gmail.com')
cf.key_write_enter('deepakchiragmali@gmail.com')
cf.key_write_enter('cmali5804@gmail.com')
cf.key_press('Tab')
cf.key_write_enter("Task-5 ")
cf.key_press('Tab')
cf.key_write_enter("Successfully Done Currency Convertor and save it into Excel")
cf.key_hit_enter()
cf.launch_any_exe_bat_application('C:/Users/chirag mali/Desktop/clointfusion task/task 5/noteToexcel.xlsx')
cf.key_press('ctrl+a')
cf.key_press('ctrl+c')
cf.window_close_windows('Excel')
cf.key_press('ctrl+v')
cf.browser_mouse_click_h("Attach")
cf.browser_mouse_click_h("Browse this computer")
time.sleep(2)
cf.key_write_enter(r'C:\Users\chirag mali\Desktop\clointfusion task\task 5\noteToexcel.xlsx')
cf.browser_mouse_click_h("Send")
# cf.browser_quit_h()
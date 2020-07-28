from tkinter import *
import time
import os
import datetime
import calendar
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.common.keys import Keys
from fuzzywuzzy import fuzz
import csv
import shutil
import pandas as pd
import chromedriver_autoinstaller

username = os.environ['USERNAME']
username = username.split(".")
name = ""
for x in username:
    x = x.capitalize()
    name = name + " " + x
username = name[1:]
userpath = os.environ['USERPROFILE']
now = datetime.datetime.now()
now = str(now.year) + " " + str(now.month) + " " + str(now.day) + " " + str(now.hour) + " " + str(now.minute)
today = datetime.date.today()
n_month = today.month - 1  # previous month as number
w_month = calendar.month_abbr[n_month]  # previous month as abbreviation of word
full_month = calendar.month_name[n_month]  # previous month full name
if n_month == 12:  # get the year of the previous month then make string
    d_year = str(today.year - 1)
else:
    d_year = str(today.year)
num_month = n_month  # integer number for month
n_month = str(n_month)
if len(n_month) == 1:  # make month a string
    n_month = "0" + str(n_month)
else:
    n_month = str(n_month)



def check_if_selected(building):
    try:
        if building in check_boxes.keys():
            if check_boxes[building] == 0:
                return False
            else:
                callback(building)
                return True
    except:
        callback(building + " is not selected.  Trying next facility")


def get_micr_name(name):  # convert pcc name to qb name
    fuzzcheck = fuzz.partial_ratio("KEYTOTAL", name)
    if fuzzcheck == 100:
        return "Match"
    else:
        return "No match"
        callback("Could not find KEYTOTAL bank")


def to_text(message):
    try:
        s = str(datetime.datetime.now().strftime("%H:%M:%S")) + ">>  " + str(message) + "\n"
        with open('P:\\PACS\\Finance\\Automation\\PCC AP Check Runs\\logs\\Py ' + username + ' ' + str(now) + '.txt', 'a') as file:
            file.write(s)
            file.close()
    except:
        pass


def write_to_csv(filename, building, date, total, entries):
    check = os.path.exists(filename)

    if check == True:
        with open(filename, 'a', newline='') as csvfile:
            filewriter = csv.writer(csvfile, delimiter=',', quoting=csv.QUOTE_MINIMAL)
            filewriter.writerow([building, date, total, entries])
    else:
        with open(filename, 'w', newline='') as csvfile:
            filewriter = csv.writer(csvfile, delimiter=',', quoting=csv.QUOTE_MINIMAL)
            filewriter.writerow(["Building", "Created Date", "Batch Total", "Num Entries"])

        with open(filename, 'a', newline='') as csvfile:
            filewriter = csv.writer(csvfile, delimiter=',', quoting=csv.QUOTE_MINIMAL)
            filewriter.writerow([building, date, total, entries])


# get latest driver from the shared drive and add to user documents folder
def find_updated_driver():
    folder = 'P:\\PACS\\Finance\\Automation\\Chromedrivers\\'
    file_list = []
    if os.path.isdir(folder):
        list_items = os.listdir(folder)
        for item in list_items:
            file = item.split(" ")
            if file[0] == 'chromedriver':
                file_list.append(file[1][:2])
        try:
            callback('Updating chrome driver to newer version')
            shutil.copyfile(folder + 'chromedriver ' + max(file_list) + '.exe',
                            os.environ['USERPROFILE'] + '\\Documents\\AP Check Runs\\chromedriver ' + max(
                                file_list) + '.exe')
        except:
            callback("Couldn't update driver automatically")
        return max(file_list)


# check the current driver version on your computer
def find_current_driver():
    folder = os.environ['USERPROFILE'] + '\\Documents\\AP Check Runs\\'
    file_list = []
    if os.path.isdir(folder):
        list_items = os.listdir(folder)
        for item in list_items:
            file = item.split(" ")
            if file[0] == 'chromedriver':
                file_list.append(file[1][:2])
        return max(file_list)


# login
class LoginPCC:
    def __init__(self):  # create an instance of this class. Begins by logging in
        try:
            # chromedriver_autoinstaller.install()
            latestdriver = find_current_driver()
            self.driver = webdriver.Chrome(
                os.environ['USERPROFILE'] + '\\Documents\\AP Check Runs\\chromedriver ' + str(latestdriver) + '.exe')
        except:
            latestdriver = find_updated_driver()
            self.driver = webdriver.Chrome(
                os.environ['USERPROFILE'] + '\\Documents\\AP Check Runs\\chromedriver ' + str(latestdriver) + '.exe')
        self.driver.get('https://login.pointclickcare.com/home/userLogin.xhtml')
        time.sleep(5)
        try:
            usernamex = self.driver.find_element(By.ID, 'username')
            usernamex.send_keys(usernametext)
            passwordx = self.driver.find_element(By.ID, 'password')
            passwordx.send_keys(passwordtext)
            self.driver.find_element(By.ID, 'login-button').click()
        except:
            self.driver.get(
                'https://www12.pointclickcare.com/home/login.jsp?ESOLGuid=40_1595946502980')  # security login page
            time.sleep(8)
            try:
                usernamex = self.driver.find_element(By.ID, 'id-un')
                usernamex.send_keys(usernametext)
                passwordx = self.driver.find_element(By.ID, 'password')
                passwordx.send_keys(passwordtext)
                self.driver.find_element(By.ID, 'id-submit').click()
            except:
                callback('There was an issue with logging into PCC.')

    def teardown_method(self):  # exit the program (FULLY WORKING)
        self.driver.quit()

    def buildingSelect(self, building):  # select your building (FULLY WORKING)
        self.driver.get("https://www30.pointclickcare.com/home/home.jsp?ESOLnewlogin=Y")
        self.driver.find_element(By.ID, "pccFacLink").click()
        time.sleep(1)
        try:
            self.driver.find_element(By.PARTIAL_LINK_TEXT, building).click()
        except:
            self.driver.get("https://www30.pointclickcare.com/home/home.jsp?ESOLnewlogin=Y")
            callback("Could not locate " + building + " in PCC")

    def Check_Run(self, checkdatetext, paythrutext):  # download the income statement m-to-m report (FULLY WORKING)
        window_before = self.driver.window_handles[0]  # make window tab object
        time.sleep(1)
        self.driver.get('https://www30.pointclickcare.com/glap/ap/processing/batchlist.jsp')
        try:
            self.driver.find_element(By.LINK_TEXT, "Payments").click()
        except:
            pass
        self.driver.find_element(By.XPATH, "/html/body/form/table/tbody/tr[5]/td/input[1]").click()
        window_after = self.driver.window_handles[1]  # set second tab
        self.driver.switch_to.window(window_after)  # select the second tab
        self.driver.find_element(By.NAME, "description").click()
        self.driver.find_element(By.NAME, "description").send_keys("Check Run " + checkdatetext)
        el = self.driver.find_element(By.NAME, 'bank_id')  # select the MICR in the dropdown menu
        for option in el.find_elements_by_tag_name('option'):
            if get_micr_name(option.text) == 'Match':
                option.click()
                break
        self.driver.find_element(By.XPATH,
                                 '//*[@id="frmData"]/table[2]/tbody/tr[5]/td[2]/input[2]').click()  # select direct payment
        self.driver.find_element(By.NAME, 'inc_only_pay').click()  # uncheck include only vendors
        self.driver.find_element(By.NAME, "check_date").click()  # update check date
        self.driver.find_element(By.NAME, "check_date").send_keys(6 * Keys.BACKSPACE)
        self.driver.find_element(By.NAME, "check_date").send_keys(6 * Keys.DELETE)
        self.driver.find_element(By.NAME, "check_date").send_keys(checkdatetext)
        dropdown = self.driver.find_element(By.NAME, "select_invoices_by")  # select Due Date
        dropdown.find_element(By.CSS_SELECTOR,
                              "#frmData > table:nth-child(14) > tbody > tr:nth-child(10) > td.data > select > option:nth-child(2)").click()  # select bank account
        self.driver.find_element(By.NAME, "on_or_before").click()
        self.driver.find_element(By.NAME, "on_or_before").send_keys(paythrutext)
        self.driver.find_element(By.XPATH,
                                 '//*[@id="frmData"]/table[2]/tbody/tr[11]/td[2]/input[4]').click()  # select vendor groups radio button
        el = self.driver.find_element(By.NAME, 'vendor_group_id')  # select the MICR in the dropdown menu
        for option in el.find_elements_by_tag_name('option'):
            if option.text == 'Providence Vendor Group':
                option.click()
                break
        time.sleep(2)
        self.driver.find_element(By.CSS_SELECTOR, "#selectButton > input").click()  # select invoices button
        window_after = self.driver.window_handles[2]  # set second window
        self.driver.switch_to.window(window_after)  # select the second window
        time.sleep(1)
        self.driver.find_element(By.NAME, "checkedAll").click()  # click check all box
        self.driver.find_element(By.CSS_SELECTOR,
                                 "body > form:nth-child(2) > table > tbody > tr:nth-child(1) > td > table > tbody > tr > td:nth-child(1) > input:nth-child(1)").click()
        window_after = self.driver.window_handles[1]  # set second window
        self.driver.switch_to.window(window_after)  # select the second window
        try:
            self.driver.find_element(By.CSS_SELECTOR, "#runButton > input").click()  # save the pmt batch
            # self.driver.close() #for testing
            self.driver.switch_to.window(window_before)  # select the original window
            createddate = self.driver.find_element(By.CSS_SELECTOR,
                                                   "body > form > table > tbody > tr:nth-child(8) > td > div > table > tbody > tr:nth-child(2) > td:nth-child(5)").text  # get text
            print(createddate)
            batchtotal = self.driver.find_element(By.CSS_SELECTOR,
                                                  "body > form > table > tbody > tr:nth-child(8) > td > div > table > tbody > tr:nth-child(2) > td:nth-child(6)").text  # get text
            batchtotal = batchtotal.replace(",", "")
            numentries = self.driver.find_element(By.CSS_SELECTOR,
                                                  "body > form > table > tbody > tr:nth-child(8) > td > div > table > tbody > tr:nth-child(2) > td:nth-child(7)").text  # get text
            # self.driver.find_element(By.CSS_SELECTOR,"body > form > table > tbody > tr:nth-child(8) > td > div > table > tbody > tr:nth-child(2) > td:nth-child(1) > a:nth-child(4)").click()  # post
        except:
            callback("Run button did not populate for this building")
            self.driver.close()
            self.driver.switch_to.window(window_before)
            createddate = 'None'
            batchtotal = 'None'
            numentries = 'None'
        write_to_csv('pmt log.csv', facname, createddate, batchtotal, numentries)

    def Check_Run_Post(self):  # download the income statement m-to-m report (FULLY WORKING)
        self.driver.get('https://www30.pointclickcare.com/glap/ap/processing/batchlist.jsp')
        try:
            self.driver.find_element(By.LINK_TEXT, "Payments").click()
        except:
            pass
        try:
            check_table = self.driver.find_element(By.CSS_SELECTOR,
                                                   "body > form > table > tbody > tr:nth-child(8) > td > div > table")  # get the table webelement
            for row in check_table.find_elements(By.CSS_SELECTOR,
                                                 "body > form > table > tbody > tr:nth-child(8) > td > div > table > tbody > tr"):  # loop through table elements
                for cell in row.find_elements(By.TAG_NAME, 'td'):  # loop through each cell of each row of the table
                    try:
                        if facname in self.driver.find_element(By.CSS_SELECTOR,
                                                               "#pccFacLink").text:  # verify the correct building is selected
                            cell_text = cell.text
                            if cell.text[:9] == "Check Run":  # only post the descriptions with check run
                                callback(facname + ': posting ' + cell_text)  # notify of the posting
                                row.find_element(By.LINK_TEXT, 'post').click()
                                alert_obj = self.driver.switch_to.alert
                                alert_obj.accept()
                                break
                    except:
                        callback("There was an issue selecting the correct building.  Going to next building")
                        break
        except:
            callback("Could not post for " + facname)


# start up selenium
def start_PCC():
    global PCC
    PCC = LoginPCC()
    time.sleep(5)


def Run_Check_Run(checkdate, paythrudate):
    global fac
    global facname
    callback("Running Check Run")
    start_PCC()
    for fac in facilitydict:
        facname = facilitydict[fac][1]
        if check_if_selected(fac):  # is this facility selected?
            PCC.buildingSelect(facname)  # go to the next building
            time.sleep(1)  # wait to load
            PCC.Check_Run(checkdate, paythrudate)  # run
    PCC.teardown_method()  # end of process
    callback("Process has finished")


def Run_Check_Run_Post():
    callback("Running Check Run Posting")
    start_PCC()
    global fac
    global facname
    for fac in facilitydict:
        facname = fac
        if check_if_selected(facname) == True:  # is this facility selected?
            PCC.buildingSelect(fac)  # go to the next building
            time.sleep(1)  # wait to load
            PCC.Check_Run_Post()  # run
    PCC.teardown_method()  # end of process
    callback("Process has finished")


def print_checkboxes():
    try:
        listoffacilities = ""
        for x in check_boxes:
            if check_boxes[x] == 1:
                listoffacilities = listoffacilities + ", " + x
        if listoffacilities == "":
            callback("No facilities have been selected")
        else:
            callback("Facilities selected: " + listoffacilities[2:])
    except:
        callback("No facilities have been selected")


# tkinter start - GUI section---------------------------------------------------------
root = Tk()  # create a GUI
root.title("Providence Group AP Payments v2020.07.29")
# root.geometry("%dx%d+%d+%d" % (1200, 400, 1000, 200))
root.resizable(False, False)


# root.iconbitmap("C:\\Users\\tyler.anderson\\Documents\\Python\\Projects\\PCC HUB\\PACS Logo.ico")


def new_winF():  # new window definition
    newwin = Toplevel(root, bg=headcolor)
    newwin.title("Select Facility")
    newwin.resizable(False, False)
    # newwin.iconbitmap("C:\\Users\\tyler.anderson\\Documents\\Python\\Projects\\PCC HUB\\PACS Logo.ico")

    boxframe = Frame(newwin, bg=headcolor, pady=10, bd=10)
    scframe = Frame(newwin, bg=headcolor, pady=4, bd=10)
    saframe = Frame(newwin, bg=headcolor, pady=4, bd=10)
    caframe = Frame(newwin, bg=headcolor, pady=4, bd=10)

    # newwin.grid_rowconfigure(0, weight=1)
    # newwin.grid_columnconfigure(0, weight=1)

    boxframe.grid(row=0, columnspan=3)
    scframe.grid(row=1, column=0)
    saframe.grid(row=1, column=1)
    caframe.grid(row=1, column=2)

    def get_value():  # get checkbox status and close window
        for status in check_boxes:
            check_boxes[status] = check_boxes[status].get()
        newwin.destroy()

    def select_all():
        for status in check_boxes:
            check_boxes[status].set(1)

    def clear_all():
        for status in check_boxes:
            check_boxes[status].set(0)

    global check_boxes
    check_boxes = {facility: IntVar() for facility in facilityindex}  # create dict of check_boxes
    i = 0
    r = 0
    for facility in facilityindex:  # loop to add boxes from list
        i += 1
        l = Checkbutton(boxframe, text=facility, variable=check_boxes[facility], bg=headcolor)
        if i <= 10:
            l.grid(row=i, column=r, sticky=W)
        else:
            r += 1
            i = 1
            l.grid(row=i, column=r, sticky=W)

    savebutton = Button(scframe, padx=2, pady=2, width=15, text="Save and Close", command=get_value)
    selectallbutton = Button(saframe, padx=2, pady=2, width=15, text="Select All", command=select_all)
    clearallbutton = Button(caframe, padx=2, pady=2, width=15, text="Clear All", command=clear_all)

    savebutton.grid(row=11, sticky="nsew")
    selectallbutton.grid(row=11, column=1, sticky="nsew")
    clearallbutton.grid(row=11, column=2, sticky="nsew")
    newwin.mainloop()


def get_date_win():  # new window definition
    newwin = Toplevel(root, bg=headcolor)
    newwin.title("Start Check Run")
    newwin.resizable(False, False)

    # newwin.iconbitmap("C:\\Users\\tyler.anderson\\Documents\\Python\\Projects\\PCC HUB\\PACS Logo.ico")

    def getentrytext():
        global usernametext
        global passwordtext
        checkdatetext = checkdateentry.get()
        paythrutext = paythrudateentry.get()
        usernametext = username.get()
        passwordtext = password.get()
        newwin.destroy()
        Run_Check_Run(checkdatetext, paythrutext)

    boxframe = Frame(newwin, bg=headcolor, pady=10, bd=10)
    scframe = Frame(newwin, bg=headcolor, pady=4, bd=10)
    saframe = Frame(newwin, bg=headcolor, pady=4, bd=10)
    caframe = Frame(newwin, bg=headcolor, pady=4, bd=10)

    # newwin.grid_rowconfigure(0, weight=1)
    # newwin.grid_columnconfigure(0, weight=1)

    boxframe.grid(row=0, columnspan=2)
    scframe.grid(row=2, column=0)
    saframe.grid(row=1, column=0)
    # welcome message
    welcomelabel = Label(boxframe,
                         text="To create System Batches in PCC, enter the check run date (example: 1/1/2019) and press 'Run'",
                         bg=headcolor)
    welcomelabel.grid(row=0, column=0)
    welcomelabel.config(font=22)
    # date labels
    checkdatelabel = Label(saframe, text="Check Date:", bg=headcolor)
    checkdatelabel.grid(row=1, column=0)
    paythrulabel = Label(saframe, text="Pay Thru Date:", bg=headcolor)
    paythrulabel.grid(row=2, column=0)
    # date entry
    checkdateentry = Entry(saframe)
    checkdateentry.grid(row=1, column=1)
    paythrudateentry = Entry(saframe)
    paythrudateentry.grid(row=2, column=1)
    # login info labels
    usernamelabel = Label(saframe, text="PCC Username", bg=headcolor)
    usernamelabel.grid(row=3, column=0)
    passwordlabel = Label(saframe, text="PCC Password", bg=headcolor)
    passwordlabel.grid(row=4, column=0)
    # login info entry
    username = Entry(saframe)
    username.insert(0, "pghc." + name[1].lower() + name.split(" ")[2].lower())
    username.grid(row=3, column=1)
    password = Entry(saframe, show="*")
    password.grid(row=4, column=1)

    runbutton = Button(scframe, padx=2, pady=2, width=15, text="Run", command=getentrytext)
    runbutton.grid(row=11, sticky="nsew")

    newwin.mainloop()


def post_checks_win():  # new window definition
    newwin = Toplevel(root, bg=headcolor)
    newwin.title("Post check run batches")
    newwin.resizable(False, False)

    # newwin.iconbitmap("C:\\Users\\tyler.anderson\\Documents\\Python\\Projects\\PCC HUB\\PACS Logo.ico")

    def getentrytext():
        global usernametext
        global passwordtext
        usernametext = username.get()
        passwordtext = password.get()
        newwin.destroy()
        Run_Check_Run_Post()

    boxframe = Frame(newwin, bg=headcolor, pady=10, bd=10)
    scframe = Frame(newwin, bg=headcolor, pady=4, bd=10)
    saframe = Frame(newwin, bg=headcolor, pady=4, bd=10)
    caframe = Frame(newwin, bg=headcolor, pady=4, bd=10)

    # newwin.grid_rowconfigure(0, weight=1)
    # newwin.grid_columnconfigure(0, weight=1)

    boxframe.grid(row=0, columnspan=2)
    scframe.grid(row=2, column=0)
    saframe.grid(row=1, column=0)
    welcomelabel.grid(row=0, column=0)
    welcomelabel.config(font=22)

    # login info labels
    usernamelabel = Label(saframe, text="PCC Username", bg=headcolor)
    usernamelabel.grid(row=3, column=0)
    passwordlabel = Label(saframe, text="PCC Password", bg=headcolor)
    passwordlabel.grid(row=4, column=0)
    # login info entry
    username = Entry(saframe)
    username.insert(0, "pghc." + name[1].lower() + name.split(" ")[2].lower())
    username.grid(row=3, column=1)
    password = Entry(saframe, show="*")
    password.grid(row=4, column=1)

    runbutton = Button(scframe, padx=2, pady=2, width=15, text="Run", command=getentrytext)
    runbutton.grid(row=11, sticky="nsew")

    newwin.mainloop()


def callback(message):  # update the statusbox gui
    s = str(datetime.datetime.now().strftime("%H:%M:%S")) + ">>" + str(message) + "\n"
    statusbox.insert(END, s)
    statusbox.see(END)
    statusbox.update()
    to_text(message)


headcolor = "#d7eef5"
framecolor = "#d7eef5"
footcolor = "#d7eef5"
statusboxcolor = "#f7f7f7"

# create the frames
headframe = Frame(root, width=400, height=300, bg=headcolor, pady=10, bd=10)
downloadsframe = Frame(root, width=400, height=400, bg=framecolor, padx=10, bd=10)
otherframe = Frame(root, width=400, height=400, bg=framecolor, padx=10, bd=10)
footframe = Frame(root, width=400, height=100, bg=footcolor, pady=10, bd=20)

# layout all of the main containters
root.grid_rowconfigure(0, weight=1)
root.grid_columnconfigure(0, weight=1)

headframe.grid(row=0, sticky="nsew", columnspan=10)
downloadsframe.grid(row=1, column=0, sticky="nsew")
otherframe.grid(row=1, column=1, sticky="nsew")
footframe.grid(row=2, sticky="nsew", columnspan=10)

statusbox = Text(footframe, height=20, width=70, bg=statusboxcolor)
statusbox.grid(row=1, column=0, columnspan=10)

# create the labels - headframe
welcomelabel = Label(headframe, text="Welcome " + username, bg=headcolor)
welcomelabel.grid(row=0, column=0)
welcomelabel.config(font=88)
currentmonthlabel = Label(headframe, text="Current Month: " + calendar.month_abbr[today.month] + " " + str(today.year),
                          bg=headcolor)
currentmonthlabel.grid(row=1, column=0, sticky="nsew")

# create the buttons
# downloads frame buttons - middle frames
checkrunbutton = Button(downloadsframe, text="Run PCC System Batch", padx=5, pady=5, width=35, command=get_date_win)
postcheckrunbutton = Button(downloadsframe, text="Post PCC Check Batches", padx=5, pady=5, width=35,
                            command=post_checks_win)
# add the buttons
checkrunbutton.grid(row=8, pady=5, sticky="nsew")
postcheckrunbutton.grid(row=9, pady=5, sticky="nsew")

# other frame buttons
choosefacbutton = Button(otherframe, text="Select Facilities", padx=5, pady=5, width=35,
                         command=new_winF)  # command linked

# add the buttons
choosefacbutton.grid(row=1, column=0, pady=5, sticky="nsew")

statuslabel = Label(footframe, text="Status Box:", bg=footcolor)
statuslabel.grid(row=0, column=0, sticky="nsew")

# get paths to map out how data flows if not connected to the VPN
try:
    faclistpath = "P:\\PACS\\Finance\\Automation\\PCC Reporting\\pcc webscraping.xlsx"
    try:
        os.mkdir(userpath + '\\Documents\\AP Check Runs\\')  # make directory for backup in documents folder
        shutil.copyfile(faclistpath, userpath + '\\Documents\\AP Check Runs\\pcc webscraping.xlsx')  # make backup file
    except FileExistsError:
        shutil.copyfile(faclistpath,
                        userpath + '\\Documents\\AP Check Runs\\pcc webscraping.xlsx')  # if folder exists just copy
except FileNotFoundError:  # if VPN is not connected use the one last saved
    try:
        faclistpath = userpath + '\\Documents\\AP Check Runs\\pcc webscraping.xlsx'
    except:
        callback("data cannot be connected. Please connect to the VPN")

# create dataframe from master listing
facility_df = pd.read_excel(faclistpath, sheet_name='Automation', index_col=0)
# remove the ALF's and ILF's
for row in facility_df.itertuples():
    if 'ALF' in row.Index or 'ILF' in row.Index:
        facility_df = facility_df.drop([row.Index])
facilityindex = facility_df.index.to_list()
fac_number = facility_df['Business Unit'].to_list()
pcc_name = facility_df['PCC Name'].to_list()
facilitydict = dict(zip(facilityindex, zip(fac_number, pcc_name)))

root.mainloop()

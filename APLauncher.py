import subprocess
import sys
import os
import shutil
import pyautogui

folder = os.environ['USERPROFILE'] + '\\Documents\\AP Check Runs\\'
if not os.path.isdir(folder):
    os.mkdir(os.environ['USERPROFILE'] + '\\Documents\\AP Check Runs\\')

file_list = []
if os.path.isdir(folder):
    list_items = os.listdir(folder)
    for item in list_items:
        file = item.split(" ")
        if file[0] == 'APCheckRuns':
            file_list.append(file[1][:6])
    try:
        current = max(file_list)
    except:
        current = 0
        print('no file')
    folder = "P:\\PACS\\Finance\\Automation\\PCC AP Check Runs\\"
    file_list = []
    if os.path.isdir(folder):
        list_items = os.listdir(folder)
        for item in list_items:
            file = item.split(" ")
            if file[0] == 'APCheckRuns':
                file_list.append(file[1][:6])
        try:
            most_current = max(file_list)
        except:
            print('no file')
        if most_current != current:  # update by downloading new file
            ask = pyautogui.confirm(text='The program has an update, do you wish to proceed?', title='Update Availible',
                              buttons=['Yes', 'No'])
            if ask == 'Yes':
                print('Updating program...Please wait.')
                shutil.copyfile(folder + 'APCheckRuns ' + max(file_list) + '.exe',
                                os.environ['USERPROFILE'] + '\\Documents\\AP Check Runs\\APCheckRuns ' + max(
                                    file_list) + '.exe')
                if current > 0:
                    os.remove(os.environ['USERPROFILE'] + '\\Documents\\AP Check Runs\\APCheckRuns ' + current + '.exe')
            else:
                pass
        try:
            subprocess.run(os.environ['USERPROFILE'] + '\\Documents\\AP Check Runs\\APCheckRuns ' + current + '.exe')
        except:
            try:
                subprocess.run(
                    os.environ['USERPROFILE'] + '\\Documents\\AP Check Runs\\APCheckRuns ' + most_current + '.exe')
            except:
                pyautogui.alert(text='You have no program installed', title='No Program', button='OK')
                sys.exit()
import os

folder = 'P:\\PACS\\Finance\\Automation\\PCC AP Check Runs\\'
file_list = []
if os.path.isdir(folder):
    list_items = os.listdir(folder)
    for item in list_items:
        file = item.split(" ")
        if file[0]=='chromedriver':
            print(file[1][:2])
            file_list.append(file[1][:2])
    print(file_list)
    print(max(file_list))

import os
import sys
import requests
import fastapi
import gui_window_module as win
import xlsxwriter
import httplib2
# from googleapiclient.discovery import build
from oauth2client.service_account import ServiceAccountCredentials
import gspread
from gspread_dataframe import get_as_dataframe, set_with_dataframe
from tkinter import messagebox
import tkinter as tk
from tkinter import Tk, ttk, Text,filedialog
import pandas as pd

CREDENTIALS=\
{
  "type": "service_account",
  "project_id": "steam-capsule-340918",
  "private_key_id": "bbcc06042a81650feaef545244344cd701a6ca2e",
  "private_key": "-----BEGIN PRIVATE KEY-----\nMIIEvwIBADANBgkqhkiG9w0BAQEFAASCBKkwggSlAgEAAoIBAQCKb/S64drUn5Hu\nOJgpXxGjhdFMtm0Okst/nFI6EDFf6Rd7UyX9VgRECgFbYVX88AI3VlUabud8Mue4\nSzzyDIeCDRPFaqviq+zwEmAIYk8LQanLAuV9TF51na6JbPi/W2WTt4Psv/jVKWMO\nPlipVLthXDaFvZ+B4RhlZ2/6kNVjcPcexGtM3k+1i6taWA1+OicZondGqUcryukD\nk15fazarHXoWsa76BaIU+GmRHzAOw50+VqYpYaEc0+AWZD0M/11LbglcstVdUGHL\nj4VJplgqLfVZaGIv3ct7Z+GVgU0Lt+hbZkgXWoWhpOlhLfRXoT+atCFq3R+jULP5\nQQq6/8W1AgMBAAECggEAMymKJDEJO1BX0dcXoR6R+cGgZv36kwq8a4Z9uxi17rDJ\n7Vl+7kAGZpDeDIQOD+MxpnrhC9pV8cgsbCdeOERaJj2oA2tXZ/fLIrRfymlutgXH\n9w+0eZiqRkSGtyJPUgU4pp2Jg4s1Lq5TffWbtcTrVWGLguTgGNN0PTG7qDozqsKP\nTC76wAJT2PoVVQylwHE0oKxB8oFbWY4cVjYBzF9RqtgB224kfFuRWetjxNb6Bbau\naYZ2QL0QICtklKQBlnLzMmJcWDV9lHxQIGwwQfX8Wd7gzw3lwkuYpaHYNRs/ebjp\naPYYUv+0JPc4zh6fIMLDuaExMeq/mQAPvF2s7KbUlwKBgQDDOETvrY2rI+jfKNnV\nVA5RxDEopYZq9fgwhuRYHAp5W/AKqbRhqDKI+4MtfeZ1tAw5k/eZsQyBf1gqbgPv\nWnD7bNxsqxlhfYiJSzSiu3OoGCiPhs+nzteL8VDdKeIQtSPiuC8Wk9HNrHX+Xahh\n12PNhM4Oy5kmXsm/rGQea537VwKBgQC1ieqRr7fYdPjV+rYTik6d9LIDQVCLGbKs\nxcd0tgkEWywNxjUFNXSRJgyNQI1FIBhusVJ1Ge1aXm+FL3RIaGu/5eCFoVcDvfAk\nFkv20i3Hu1vdi3S8QiLrMnv1T82sFHNx+28y0BgoE+0eocJodNuVDlj/Vjmlp5FJ\n8Nx3LL4r0wKBgQCV11faW+00YjC5MULaIlWHXz6YQ0zERp3EqZUVKBjGA9JgbNfe\naVq4l8ydG1jMGXGUtEVFO4cs0pDaqGzuyA2Wfn1GD6JAmTk2oHn7OkRQzpI7cC9t\nTy9U49m8mAxD5LVxrQu/maBc7LX4kuzOhKO/OONsqcuYjwLt0yVZ0CKHqwKBgQCx\nxYGn8qwU0s8OM7nzPqAn/AQKPf6SiLK4j+D3AH+p/WIRhwRKuoMQ1HK8K/drNrfW\nRdzagW417X5FrSew9Fh3jbOlCE5+gpRTsmXnKQDdszKNq8+/vwAU09YhbmmY1loK\nx06oMrFFJeYw9fS7d5vDxk4OlLBU8NfM0YoDRhRgMwKBgQDC4t7Jhhtk60ozBFwm\n/WvmZB5d42lEKeLZdPTT4eF5kkKroIyozm9m4/HmhAJkHqxndbkrhOu5ePIYgj78\nR94x8aGsFk5L1zOIPDXCWRDrIfkO5J4/8+yMjTBuPDpWxuEpo2dF4JGVRydSDgRz\nxow61okKAdmXUrbeg/7XcX8/4Q==\n-----END PRIVATE KEY-----\n",
  "client_email": "odessa11@steam-capsule-340918.iam.gserviceaccount.com",
  "client_id": "105357641307897168369",
  "auth_uri": "https://accounts.google.com/o/oauth2/auth",
  "token_uri": "https://oauth2.googleapis.com/token",
  "auth_provider_x509_cert_url": "https://www.googleapis.com/oauth2/v1/certs",
  "client_x509_cert_url": "https://www.googleapis.com/robot/v1/metadata/x509/odessa11%40steam-capsule-340918.iam.gserviceaccount.com"
}

# def get_service_simple():
#    return build('sheets', 'v4', developerKey=creds.api_key)


# def get_service_sacc():
#     """
#     Могу читать и (возможно) писать в таблицы кот. выдан доступ
#     для сервисного аккаунта приложения
#     sacc-1@privet-yotube-azzrael-code.iam.gserviceaccount.com
#     :return:
#     """
#     creds_json = os.path.dirname(__file__) + "/creds/sacc1.json"
#     scopes = ['https://www.googleapis.com/auth/spreadsheets']

#     creds_service = ServiceAccountCredentials.from_json_keyfile_name(creds_json, scopes).authorize(httplib2.Http())
#     return build('sheets', 'v4', http=creds_service)


# # service = get_service_simple()
# service = get_service_sacc()
# sheet = service.spreadsheets()

# # https://docs.google.com/spreadsheets/d/xxx/edit#gid=0
# sheet_id = "1Pt23r8uW0oc5FjdyF0Vcn0c4kgJmN7-5YNzgjJQYZ4c"

# https://developers.google.com/sheets/api/reference/rest/v4/spreadsheets.values/get
# resp = sheet.values().get(spreadsheetId=sheet_id, range="Лист1!A1:A999").execute()

# https://developers.google.com/sheets/api/reference/rest/v4/spreadsheets.values/batchGet

# No gspread google client
# creds_json = os.path.dirname(__file__) + "/creds/sacc1.json"
# gc = gspread.service_account(filename= creds_json)
# # sheet_id = '1Pt23r8uW0oc5FjdyF0Vcn0c4kgJmN7-5YNzgjJQYZ4c'
# sht= gc.open_by_url('https://docs.google.com/spreadsheets/d/1Pt23r8uW0oc5FjdyF0Vcn0c4kgJmN7-5YNzgjJQYZ4c/edit#gid=0')
#worksheet=sht.sheet1
# resp = sht.values().batchGet(spreadsheetId=sheet_id, ranges=["Sheet1"]).execute()
def file_open_dialog():
            global file_name; global path_file; global str_data; global n
            # Build a list of tuples for each file type the file dialog should display
            my_filetypes = [('excel file','xlsx')]
            # Ask the user to select a folder.
            # path_file = filedialog.askdirectory(parent=win,
            #                               initialdir=os.getcwd(),
            #                               title="Please select a folder:")

            # Ask the user to select a single file name.
            file_name = filedialog.askopenfilename(parent=wind,
                                    initialdir=os.getcwd(),
                                    title="Please select a file:",
                                    filetypes=my_filetypes)
            if file_name=='':
                messagebox.showinfo("Information","File not selected! \n Please select file")
            return file_name


def make_cooglesheets_client():
    global ggl_wrksheet
    # READ CREDIDS FROM FILE
    # creds_json = os.path.dirname(__file__) + "/creds/sacc1.json"
    # gc = gspread.service_account(filename= creds_json)
    # sheet_id = '1Pt23r8uW0oc5FjdyF0Vcn0c4kgJmN7-5YNzgjJQYZ4c'

    gc = gspread.service_account_from_dict(CREDENTIALS)
    str_url=read_data_textbox()
    if str_url=='':
        return False
    else:
        sht= gc.open_by_url(str_url)
        # CHOICE SHEET IN WORKBOOK !!!!
        ggl_wrksheet_list=sht.worksheets()
        ggl_wrksheet=sht.get_worksheet(0)
        # ggl_wrksheet=sht.worksheet('זמני עבודה')
        return True

def read_from_googlesheets():
    global str_data; global lst_result_ggl; global dataframe
    global ggl_wrksheet
    make_cooglesheets_client()
    if make_cooglesheets_client()==True:
        lst_result_ggl=ggl_wrksheet.get_all_values()
        dataframe=pd.DataFrame(ggl_wrksheet.get_all_records())
    else:
        messagebox.showinfo("Information","URL is empty ! \n Please input URL")

def write_to_googlesheets():
    global str_data; global lst_result_ggl; global dataframe; global ggl_wrksheet
    make_cooglesheets_client()
    lst_result_ggl=ggl_wrksheet.get_all_values()
    df = get_as_dataframe(ggl_wrksheet, parse_dates=True, header=None)
    set_with_dataframe(ggl_wrksheet,dataframe)
    # ggl_wrksheet.update(dataframe)
    # ggl_wrksheet.update('A1:B2', [[1, 2], [3, 4]])
    # =pd.Dataframe(ggl_wrksheet.get_all_records())
    # print(row_number,'\n', column_number)
def write_to_file():
    # try:
       file_name= os.path.dirname(__file__) + '/'+'new_file_name.txt'
       data= open(file_name,'w')
    #    data=open('C:\\Users\\migmcher\\Programming\\Projects\\Google_API\\new_file_name.txt','w')
    #    str_values=str(read_from_googlesheets())
       read_from_googlesheets()
       data.write(str(lst_result_ggl))
       data.close()
        # messagebox.showinfo("Information","File created and writed ! \n Succefuly")
    # except:
        # messagebox.showinfo("Information")
        # raise TypeError('File not created!')
def read_frome_txt_file():
    global lst_result_ggl
    # try:
    file_name= os.path.dirname(__file__) + '/'+'new_file_name.txt'
    data= open(file_name,'r')
        #   data=open('C:\\Users\\migmcher\\Programming\\Projects\\Google_API\\new_file_name.txt','w')
        #    str_values=str(read_from_googlesheets())
    lst_result_ggl=data
    print(lst_result_ggl)
    data.close()
        # messagebox.showinfo("Information","File created and writed ! \n Succefuly")
    # except:
        # messagebox.showinfo("Information")
        # raise TypeError('File not created!')
def read_from_excel_file():
    global lst_result_ggl; global ggl_wrksheet; global dataframe
    # try:
    file_name=file_open_dialog()
    dataframe=pd.read_excel(file_name)
    # row_dim,col_dim=make_cooglesheets_client()


        #   data=open('C:\\Users\\migmcher\\Programming\\Projects\\Google_API\\new_file_name.txt','w')
        #    str_values=str(read_from_googlesheets())
    # lst_result_ggl=data
    # print(lst_result_ggl)
    # data.close()
def read_data_textbox():
    global str_data
    str_data = str(wind.txt_text.get('1.0','end-1 c')) #!!!!!!! end-1 c - remove addition char from your INPUT !!!!!
    return str_data
def write_to_excel():
    global lst_result_ggl
    workbook = xlsxwriter.Workbook('result.xlsx')
    worksheet = workbook.add_worksheet('Sheet2')
    i=0
    print(lst_result_ggl)
    while i <len(lst_result_ggl):
        ind_sublist=0
        while ind_sublist<len(lst_result_ggl[i]):
            worksheet.write(i,ind_sublist, lst_result_ggl[i][ind_sublist])
            ind_sublist+=1
        i+=1
    workbook.close()

if __name__== '__main__':
    file_name=''; ggl_wrksheet=''; path_file=''; str_data=''; lst_result_ggl=''
    dataframe=pd.DataFrame()

    wind=win.Gui(350,400,10,10); wind.title('Good')
    # wind.bt_second.destroy()
    wind.bt_first['text']='Read from \n google table'; wind.bt_first['command']=read_from_googlesheets
    wind.bt_second['text']='Write to file'; wind.bt_second['command']=write_to_excel
    wind.bt_third['text']='Read from \n excel'; wind.bt_third['command']=read_from_excel_file
    wind.bt_four['text']='Write to Google \n Table'; wind.bt_four['command']=write_to_googlesheets
    wind.txt_text['width']=40; wind.txt_text.grid(columnspan=3)
    wind.lbl_name['text']='URL of google Table'; wind.lbl_name['width']=20

    # write_to_file(str(read_from_googlesheets))
    # print(values_list)
    wind.mainloop()
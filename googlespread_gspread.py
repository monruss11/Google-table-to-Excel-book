import os
import sys
import requests
import gui_window_module as win
import xlsxwriter
import httplib2
from oauth2client.service_account import ServiceAccountCredentials
import gspread
from gspread_dataframe import get_as_dataframe, set_with_dataframe
from tkinter import messagebox
import tkinter as tk
from tkinter import Tk, ttk, Text,filedialog
import pandas as pd


def make_googlesheets_client():
  global ggl_wrksheet
  # READ CREDIDS FROM FILE
  creds_json = os.path.dirname(__file__) + '/creds/credentials.json'
  gc = gspread.service_account(filename= creds_json)
  str_url=read_data_textbox()
  if str_url=='':
    return False
  else:
    try:
      sht= gc.open_by_url(str_url)
      gspread.exceptions.NoValidUrlKeyFound()
    except:
      return False
    # CHOICE SHEET IN WORKBOOK !!!!
    ggl_wrksheet=sht.get_worksheet(0)
    return True

def read_from_googlesheets():
  global str_data; global lst_result_ggl; global dataframe
  global ggl_wrksheet
  if make_googlesheets_client()==True:
    lst_result_ggl=ggl_wrksheet.get_all_values()
    dataframe=pd.DataFrame(ggl_wrksheet.get_all_records())
  else:
    messagebox.showinfo('Information','URL is empty or not valid! \n Please input URL')

def write_to_googlesheets():
  global str_data; global lst_result_ggl; global dataframe; global ggl_wrksheet
  make_googlesheets_client()
  lst_result_ggl=ggl_wrksheet.get_all_values()
  df = get_as_dataframe(ggl_wrksheet, parse_dates=True, header=None)
  set_with_dataframe(ggl_wrksheet,dataframe)

def write_to_file():
  try:
    file_name= os.path.dirname(__file__) + '/'+'new_file_name.txt'
    data= open(file_name,'w')
    read_from_googlesheets()
    data.write(str(lst_result_ggl))
    data.close()
    messagebox.showinfo('Information','File created and writed ! \n Succefuly')
  except:
    messagebox.showinfo('Information')
    raise TypeError('File not created!')

def read_frome_txt_file():
  global lst_result_ggl
  file_name= os.path.dirname(__file__) + '/'+'new_file_name.txt'
  data= open(file_name,'r')
  lst_result_ggl=data
  data.close()

def read_from_excel_file():
  global lst_result_ggl; global ggl_wrksheet; global dataframe
  # try:
  file_name=file_open_dialog()
  dataframe=pd.read_excel(file_name)

def read_data_textbox():
  global str_data
  str_data = str(wind.txt_text.get('1.0','end-1 c')) #!!!!!!! end-1 c - remove addition char from your INPUT !!!!!
  return str_data

def write_to_excel():
  global lst_result_ggl
  workbook = xlsxwriter.Workbook('result.xlsx')
  worksheet = workbook.add_worksheet('Sheet2')
  i=0
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

  wind.bt_first['text']='Read from \n google table'; wind.bt_first['command']=read_from_googlesheets
  wind.bt_second['text']='Write to \n Excel file'; wind.bt_second['command']=write_to_excel
  wind.bt_third['text']='Read from \n excel'; wind.bt_third['command']=read_from_excel_file
  wind.bt_four['text']='Write to Google \n Table'; wind.bt_four['command']=write_to_googlesheets
  wind.txt_text['width']=40; wind.txt_text.grid(columnspan=3)
  wind.lbl_name['text']='URL of google Table'; wind.lbl_name['width']=20

  wind.mainloop()

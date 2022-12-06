import shutil
import tkinter.ttk
import threading
import requests
import win32com.client
import os
import glob
from tkinter import *


WHITE = "#ffffff"
alpine_bi_url = "https://*********.com/Actions/UserAction.php"


params = {
    "call": "loginUser",
    "username": '************@*******.com',
    "password": '*********'
}

download_urls = {
 "file_to_downloads": "curl of files to download"
}


window = Tk()
window.title("Alpine BI Downloader")
window.config(padx=50, pady=50, bg=WHITE)

pb = tkinter.ttk.Progressbar(
    window,
    orient='horizontal',
    mode='determinate',
    length=120
)
pb.grid(column=1, row=3, pady=20)

r = 120/len(download_urls)


def download():
    pb["value"] = 0
    with requests.Session() as s:  # create persistent session for cookies to be used in downloading download_urls
        s.get(alpine_bi_url, params=params)
        for url in download_urls:
            x = s.get(url= download_urls[url],stream=True)
            with open(f'download/{url}.xls', 'wb') as out_file:
                shutil.copyfileobj(x.raw, out_file)
            pb["value"] += r
            window.update_idletasks()


def convert():
    pb["value"] = 0
    excel = win32com.client.Dispatch("Excel.Application")
    excel.visible = False

    input_dir = os.path.abspath("./download")
    output_dir = os.path.abspath("./converted")
    files = glob.glob(input_dir + "/*.xls")  # returns file objects of matching files from input_dir
    excel.DisplayAlerts = False

    for filename in files:
        file = os.path.basename(filename)  # extract filename
        output = output_dir + '/' + file.replace('.xls', '.xlsx')
        wb = excel.Workbooks.Open(filename)  # open xls files from input directory
        wb.ActiveSheet.SaveAs(output,51)
        wb.Close(True)
        pb["value"] += r
        window.update_idletasks()


canvas = Canvas(width=200, height=200, bg=WHITE, highlightthickness=0)
logo_img = PhotoImage(file="images/logo.png")
canvas.create_image(100, 100, image=logo_img)
canvas.grid(column=1, row=0)

download_button = Button(text="Download", font=("Arial", 8, "normal"), width=17, bg=WHITE,
                         command=threading.Thread(target=download).start)
download_button.grid(column=1, row=1)

convert_button = Button(text="Convert", font=("Arial", 8, "normal"), width=17, bg=WHITE, command=convert)
convert_button.grid(column=1, row=2)


window.mainloop()

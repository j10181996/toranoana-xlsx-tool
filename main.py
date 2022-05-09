#!/usr/bin/python
# -*- coding: utf-8 -*-
import time
import io
import threading
import tkinter as tk
from tkinter import ttk, messagebox
from os.path import exists
import requests
from bs4 import BeautifulSoup
from openpyxl import Workbook, load_workbook
from openpyxl.drawing.image import Image
from openpyxl.styles import Alignment
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.support.select import Select

class App(tk.Tk):
    def __init__(self):
        super().__init__()

        self.title('Toranoana xlsx tool')
        self.face_load = tk.Frame(self)
        self.progressbar = ttk.Progressbar(self.face_load, length=280, mode="indeterminate")
        self.progressbar.grid(column=0, row=0, padx=10, pady=10)

        self.face_ipt = tk.Frame(self, width=300)
        self.face_ipt.pack()
        email_label = tk.Label(self.face_ipt, text='Account:')
        password_label = tk.Label(self.face_ipt, text='Password:')
        path_label = tk.Label(self.face_ipt, text='Path')
        self.email = tk.Entry(self.face_ipt)
        self.password = tk.Entry(self.face_ipt, show="*")
        self.path = tk.Entry(self.face_ipt)
        button = tk.Button(self.face_ipt, text='Start', command=self.start)
        email_label.grid(column=0, row=0, padx=10, pady=10)
        self.email.grid(column=1, row=0, padx=10, pady=10)
        password_label.grid(column=0, row=1, padx=10, pady=10)
        self.password.grid(column=1, row=1, padx=10, pady=10)
        path_label.grid(column=0, row=2, padx=10, pady=10)
        self.path.grid(column=1, row=2, padx=10, pady=10)
        button.grid(column=2, row=1, padx=10, pady=10)
        
    def start(self):
        email = self.email.get()
        password = self.password.get()
        path = self.path.get()
        if path[-5:] != ".xlsx":
            path += '.xlsx'
        self.face_ipt.pack_forget()
        self.face_load.pack()
        self.progressbar.start()
        t = ToranoanaXlsxTool(email, password, path)
        thread = threading.Thread(target=t.main)
        thread.start()
        self.monitor(thread)

    def monitor(self, thread):
        if thread.is_alive():
            self.after(100, lambda: self.monitor(thread))
        else:
            self.progressbar.stop()
            self.face_load.pack_forget()
            self.face_ipt.pack()
            messagebox.showinfo("Message",  "Done!")

class ToranoanaXlsxTool:
    def __init__(self, email, password, path):
        self.path = path if path else 'toranoana.xlsx'
        file_exists = exists(self.path)
        self.wb = load_workbook(self.path) if file_exists else Workbook()
        self.email = email
        self.password = password

    def getSheet(self, genre):
        if genre in self.wb: return self.wb[genre]
        else: 
            sheet = self.wb.create_sheet(genre)
            titles = ["封面", "CP", "社團", "作者", "書名", "價錢（日幣）", "價錢（台幣）", "出版日", "網址"]
            sheet.column_dimensions["A"].width = 100 / 7
            sheet.column_dimensions["F"].width = 16
            sheet.column_dimensions["G"].width = 16
            sheet.column_dimensions["H"].width = 16
            sheet.column_dimensions["I"].width = 55
            for row in sheet["A1":"I1"]:
                for i, cell in enumerate(row):
                    cell.value = titles[i]
                    cell.alignment = Alignment(horizontal='center', vertical='center')
            return sheet

    def moveImage(self, sheet, index):
        sheet_images = sheet._images
        for image in sheet_images:
            old_idx = 0
            if type(image.anchor) == str:
                old_idx = int(image.anchor[1:])
            else:
                old_idx = image.anchor._from.row + 1
            if old_idx >= index:
                image.anchor = "A"+str(old_idx+1)
                self.resizeImage(image)
                sheet.row_dimensions[old_idx+1].height = sheet.row_dimensions[old_idx].height

    def resizeImage(self, image):
        scale = image.height / image.width
        image.width = 100
        image.height = 100 * scale

    def getRowIndex(self, sheet, circle, author, title, date):
        r, insert_index = 1, 0
        for row in sheet.iter_rows(min_col=3, max_col=8, values_only=True):
            if row[0] == circle and row[1] == author:
                if row[2] == title:
                    return 0
                if date < row[5]:
                    sheet.insert_rows(r)
                    self.moveImage(sheet, r)
                    return r
                else:
                    insert_index = r+1
            r += 1
        if insert_index:
            sheet.insert_rows(insert_index)
            self.moveImage(sheet, insert_index)
            return insert_index
        return sheet.max_row + 1

    def login(self):
        url = 'https://ecs.toranoana.jp/ec/app/common/login/'
        self.driver.get(url)
        time.sleep(5)
        email = self.driver.find_element(By.ID, 'email')
        email.send_keys(self.email)
        password = self.driver.find_element(By.ID, 'repass')
        password.send_keys(self.password)
        button = self.driver.find_element(By.ID, 'submitLoginButton')
        button.click()

    def download(self):
        url = 'https://ecs.toranoana.jp/ec/app/mypage/order_history/'
        response = self.session.get(url)
        soup = BeautifulSoup(response.text, "html.parser")
        pager = soup.select_one("#pager")
        if not pager:
            return
        max_page = int(pager.get("data-maxpage"))

        for page in range(1, max_page+1):
            url = 'https://ecs.toranoana.jp/ec/app/mypage/order_history/?&currentPage=' + str(page)
            response = self.session.get(url)
            soup = BeautifulSoup(response.text, "html.parser")
            orders = soup.select(".hist-table4-information-title")

            for n in orders:
                url = n.select_one("a").get("href")
                res = requests.get(url)
                if res.status_code != 200:
                    continue
                soup = BeautifulSoup(res.text, "html.parser")
                image_url = soup.select_one(".product-detail-image-main img").get("src")
                image_res = requests.get(image_url)
                image_file = io.BytesIO(image_res.content)
                image = Image(image_file)
                self.resizeImage(image)
                genre = couple = date = ""
                spec_table = soup.select(".product-detail-spec-table tr")
                for subsouop in spec_table:
                    t = subsouop.select_one("td").get_text()
                    if "ジャンル" in t:
                        genre = subsouop.select("td")[1].select_one("a span").get_text().replace("/", " ")
                    elif t == "カップリング":
                        couple = subsouop.select("td")[1].select_one("a span").get_text()
                    elif "発行日" in t:
                        date = subsouop.select("td")[1].select_one("a span").get_text()
                title = soup.select_one(".product-detail-desc-title").get_text()
                print(title)
                if not soup.select_one(".sub-circle .sub-p span"):
                    continue
                circle = soup.select_one(".sub-circle .sub-p span").get_text()
                author = soup.select_one(".sub-name .sub-p a").get_text()
                price = int(soup.select_one(".pricearea__price").get_text().replace("円 （税込） ", "").replace(",", ""))

                sheet = self.getSheet(genre)

                values = [couple, circle, author, title, price, "", date, url]
                r = self.getRowIndex(sheet, circle, author, title, date)
                if r == 0:
                    continue
                sheet.add_image(image, 'A' + str(r))
                for row in sheet.iter_rows(min_row=r, min_col=2, max_row=r, max_col=len(values)+1):
                    for i, cell in enumerate(row):
                        cell.value = values[i]
                        if i == 7:
                            cell.hyperlink = values[i]
                            cell.style = "Hyperlink"
                        if i < 4:
                            sheet.column_dimensions[cell.column_letter].width = max(sheet.column_dimensions[cell.column_letter].width, len(values[i])*2.5)
                        cell.alignment = Alignment(horizontal='center', vertical='center')
                sheet.row_dimensions[r].height = image.height * 3 / 4
                time.sleep(5)
        if "Sheet" in self.wb: del self.wb["Sheet"]
        self.wb.save(self.path)

    def main(self):
        self.driver = webdriver.Chrome()
        self.driver.minimize_window()
        self.login()
        time.sleep(3)
        cookies = self.driver.get_cookies()
        self.session = requests.Session()
        for cookie in cookies:
            self.session.cookies.set(cookie['name'], cookie['value'])
        self.session.headers.update({ 'User-Agent': 'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/100.0.4896.127 Safari/537.36' })
        self.driver.get('https://ecs.toranoana.jp/ec/app/mypage/order_history/')
        ship_list = ['not_shipped', 'shipped']
        period_list = ['this_year', 'last_year', 'two_years_ago']

        for s in ship_list:
            radio = self.driver.find_element(By.ID, s).find_element(By.XPATH, '..')
            radio.click()
            for p in period_list:
                select = Select(self.driver.find_element(By.ID, 'searchPeriod'))
                button = self.driver.find_element(By.ID, 'submit')
                select.select_by_value(p)
                button.click()
                time.sleep(5)
                self.download()

app = App()
app.mainloop()
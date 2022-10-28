import sys, os
from PyQt5 import uic
from PyQt5.QtGui import *
from PyQt5.QtWidgets import *
from PyQt5.QtCore import *
from bs4 import BeautifulSoup
import requests
import openpyxl
import time

from urllib3 import Timeout


def resource_path(relative_path):
    """ Get absolute path to resource, works for dev and for PyInstaller """
    base_path = getattr(sys, '_MEIPASS', os.path.dirname(os.path.abspath(__file__)))
    return os.path.join(base_path, relative_path)

icon = resource_path('assets/memo.ico')
form = resource_path('ui/main.ui')

form_class = uic.loadUiType(form)[0]

filename = 'nickname.xlsx'

def worldCheck(nickname):
    try:
        nowURL = "https://maplestory.nexon.com/Ranking/World/Total?c="+nickname+"&w=0"
        raw = requests.get(nowURL,headers={'User-Agent':'Mozilla/5.0'})
        html = BeautifulSoup(raw.text,"html.parser")
        valid = html.select_one('tr.search_com_chk')
        valid.select_one('dl > dt > a').text
        return "true"
            
    except AttributeError:
            return "false"

def rebootCheck(nickname):
    try:
        nowURL2 = "https://maplestory.nexon.com/Ranking/World/Total?c="+nickname+"&w=254"
        raw2 = requests.get(nowURL2,headers={'User-Agent':'Mozilla/5.0'})
        html2 = BeautifulSoup(raw2.text,"html.parser")
        valid2 = html2.select_one('tr.search_com_chk')
        valid2.select_one('dl > dt > a').text
        return 'true'
    except AttributeError:
        return "false"



class WindowClass(QMainWindow, form_class):

    def __init__(self):
        super().__init__()
        self.setupUi(self)

        #프로그램 기본설정
        self.setWindowIcon(QIcon(icon))
        self.setWindowTitle('Nickname Checker')
        self.statusBar().showMessage('프로그램 정상 구동 중')

        #실행 후 기본값 설정

        #버튼 기능
        self.btn_start.clicked.connect(self.main)
        self.btn_exit.clicked.connect(self.exit)
        self.input_nickname.returnPressed.connect(self.main)
        self.btn_check.clicked.connect(self.main2)
        self.btn_save.clicked.connect(self.save)

        self.input_nickname.setFocus()

    
    def main(self):
        nickname = self.input_nickname.text()
        if nickname != "":
            worldChecked = worldCheck(nickname)
            time.sleep(1)
            rebootChecked = rebootCheck(nickname)
            
            print(worldChecked)
            print(rebootChecked)
            

            if worldChecked == "false" and rebootChecked == "false":
                self.label_nickname.setText(nickname)
                self.label_valid.setText("블추 시도 가능")
                self.label_valid.setStyleSheet("Color: Green")
                self.label_nickname.setStyleSheet("Color: Green")
                print('both none')

            else:
                self.label_nickname.setText(nickname)
                self.label_valid.setText('사용 중인 닉네임')
                self.label_valid.setStyleSheet("Color: Red")
                self.label_nickname.setStyleSheet("Color: Red")
                print('else')
        
        else:
            self.statusBar().showMessage('닉네임을 입력해주세요')

    def main2(self):
        global validlist
        count = 0
        validCount = 0
        validlist = []
        invalidlist = []
        try:
            data = openpyxl.load_workbook(filename)  
            sheet = data.active

            for i in list(sheet.columns)[0]:
                count += 1
                worldChecked = worldCheck(i.value)
                time.sleep(1)
                rebootChecked = rebootCheck(i.value)

                
                print(i.value)
                print(worldChecked)
                print(rebootChecked)
                
                if worldChecked == "false" and rebootChecked == "false":
                    validlist.append(i.value)
                    validCount += 1
                    self.validList.append(i.value)
                
                else:
                    invalidlist.append(i.value)
    
            self.validCount.setText(str(validCount)+" 개")
            self.nicknameCount.setText(str(count)+" 개")

        except FileNotFoundError:
            self.statusBar().showMessage('파일이 존재하지 않습니다. nickname.xlsx')


    def save(self):
        
        wb = openpyxl.Workbook()
        ws1 = wb.active
        for i in range(len(validlist)):
            ws1.append([validlist[i]])
        new_filename = 'blacklist.xlsx'
        wb.save(new_filename)
        self.statusBar().showMessage('블추 시도 가능한 닉네임이 저장됐습니다. blacklist.xlsx')

    def exit(self):
        sys.exit(0)


if __name__ == "__main__":
    app = QApplication(sys.argv) 
    myWindow = WindowClass() 
    myWindow.show()
    app.exec_()
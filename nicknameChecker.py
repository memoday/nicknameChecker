import sys, os
from PyQt5 import uic
from PyQt5.QtGui import *
from PyQt5.QtWidgets import *
from PyQt5.QtCore import *
from bs4 import BeautifulSoup
import requests
import openpyxl
import time
from fake_useragent import UserAgent

ua = UserAgent()
print(ua.chrome)

__version__ = 'v1.1.1'

latest_url = "https://api.github.com/repos/memoday/nicknameChecker/releases/latest"
gitAPI = requests.get(latest_url).json()
print('Now version: '+__version__)
print('Latest Version: '+gitAPI['tag_name'])
__latest_version__ = gitAPI['tag_name']

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
        raw = requests.get(nowURL,headers={'User-Agent':str(ua.chrome)})
        html = BeautifulSoup(raw.text,"html.parser")
        valid = html.select_one('tr.search_com_chk')
        valid.select_one('dl > dt > a').text
        return "true"

    except AttributeError:
            return "false"

def rebootCheck(nickname):
    try:
        nowURL2 = "https://maplestory.nexon.com/Ranking/World/Total?c="+nickname+"&w=254"
        raw2 = requests.get(nowURL2,headers={'User-Agent':str(ua.chrome)})
        html2 = BeautifulSoup(raw2.text,"html.parser")
        valid2 = html2.select_one('tr.search_com_chk')
        valid2.select_one('dl > dt > a').text
        return 'true'
    except AttributeError:
        return "false"

class check(QThread):
    def __init__(self, parent):
        super().__init__(parent)
        self.parent = parent
    
    def run(self):
        print('check run')
        self.parent.btn_start.setDisabled(True)
        nickname = self.parent.input_nickname.text()
        self.parent.statusBar().showMessage('????????? ?????? ???..'+nickname)
        try:
            if nickname != "":
                worldChecked = worldCheck(nickname)
                time.sleep(1)
                rebootChecked = rebootCheck(nickname)
                
                print(worldChecked)
                print(rebootChecked)
                

                if worldChecked == "false" and rebootChecked == "false":
                    self.parent.label_nickname.setText(nickname)
                    self.parent.btn_start.setEnabled(True)
                    self.parent.label_valid.setText("?????? ?????? ??????")
                    self.parent.label_valid.setStyleSheet("Color: Green")
                    self.parent.label_nickname.setStyleSheet("Color: Green")
                    print('both none')

                else:
                    self.parent.label_nickname.setText(nickname)
                    self.parent.btn_start.setEnabled(True)
                    self.parent.label_valid.setText('?????? ?????? ?????????')
                    self.parent.label_valid.setStyleSheet("Color: Red")
                    self.parent.label_nickname.setStyleSheet("Color: Red")
                    print('else')
                self.parent.statusBar().showMessage('???????????? ?????? ?????? ???')
            
            else:
                self.parent.btn_start.setEnabled(True)
                self.parent.statusBar().showMessage('???????????? ??????????????????')
        except TimeoutError:
            self.parent.statusBar().showMessage('?????? ????????? ?????? ????????????. ?????? ?????? ?????? ??????????????????.')
        
class checkList(QThread):
    def __init__(self, parent):
        super().__init__(parent)
        self.parent = parent
    
    def run(self):
        self.parent.btn_check.setDisabled(True)
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
                time.sleep(1.3)
                rebootChecked = rebootCheck(i.value)

                print(i.value)
                print(worldChecked)
                print(rebootChecked)
                self.parent.statusBar().showMessage(i.value)
                
                if worldChecked == "false" and rebootChecked == "false":
                    validlist.append(i.value)
                    validCount += 1
                    self.parent.validList.append(i.value)
                
                else:
                    self.parent.statusBar().showMessage('????????? ?????? ???..'+i.value)
            
            self.parent.statusBar().showMessage('???????????? ?????? ?????? ???')
            self.parent.btn_check.setEnabled(True)
    
            self.parent.validCount.setText(str(validCount)+" ???")
            self.parent.nicknameCount.setText(str(count)+" ???")

        except FileNotFoundError:
            self.parent.btn_check.setEnabled(True)
            self.parent.statusBar().showMessage('????????? ???????????? ????????????. nickname.xlsx')

class WindowClass(QMainWindow, form_class):

    def __init__(self):
        super().__init__()
        self.setupUi(self)

        self.label_version.setText('???????????? '+__version__)
        self.label_latestVersion.setText('???????????? '+__latest_version__)

        #???????????? ????????????
        self.setWindowIcon(QIcon(icon))
        self.setWindowTitle('Nickname Checker')
        self.statusBar().showMessage('???????????? ?????? ?????? ???')

        #?????? ??? ????????? ??????

        #?????? ??????
        self.btn_start.clicked.connect(self.main)
        self.btn_exit.clicked.connect(self.exit)
        self.input_nickname.returnPressed.connect(self.main)
        self.btn_check.clicked.connect(self.main2)
        self.btn_save.clicked.connect(self.save)

        self.input_nickname.setFocus()

    def main(self):
        print('main')
        x = check(self)
        x.start()

    def main2(self):
        print('main2')
        x = checkList(self)
        x.start()
        self.btn_check.setEnabled(True)

    def save(self):
        
        wb = openpyxl.Workbook()
        ws1 = wb.active
        for i in range(len(validlist)):
            ws1.append([validlist[i]])
        new_filename = 'blacklist.xlsx'
        wb.save(new_filename)
        self.statusBar().showMessage('?????? ?????? ????????? ???????????? ??????????????????. blacklist.xlsx')

    def exit(self):
        sys.exit(0)


if __name__ == "__main__":
    app = QApplication(sys.argv) 
    myWindow = WindowClass() 
    myWindow.show()
    app.exec_()
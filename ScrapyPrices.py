# Copyright (c) 2025 Vadym [https://github.com/Vadym-Kozichev]
# This software is distributed under the MIT License.
# For details, see the LICENSE file or visit https://opensource.org/licenses/MIT

"""
DISCLAIMER:
This project was created for educational purposes only.
It demonstrates techniques related to web scraping and browser automation using Selenium and PyQt6.
This code includes measures that may affect or bypass bot protection (e.g., Cloudflare) purely for academic interest.
The author does not condone or encourage unauthorized data scraping or violation of websites’ terms of service.
Use responsibly and ensure compliance with applicable laws and website rules.
"""

import sys
import pandas as pd
from tkinter import filedialog as fd
from tkinter.messagebox import showerror
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from PyQt6 import QtCore, QtGui, QtWidgets
from PyQt6.QtCore import QRunnable, QThreadPool, pyqtSignal, QObject


# ==== Interface ====
class Ui_MainWindow(object):
    
    def setupUi(self, MainWindow):
        MainWindow.setObjectName("MainWindow")
        MainWindow.resize(500, 250)
        self.centralwidget = QtWidgets.QWidget(parent=MainWindow)
        self.centralwidget.setObjectName("centralwidget")
        self.verticalLayout = QtWidgets.QVBoxLayout(self.centralwidget)
        self.verticalLayout.setObjectName("verticalLayout")
        self.layoutFile = QtWidgets.QHBoxLayout()
        self.layoutFile.setObjectName("layoutFile")
        self.lineEdit_file = QtWidgets.QLineEdit(parent=self.centralwidget)
        self.lineEdit_file.setObjectName("lineEdit_file")
        self.layoutFile.addWidget(self.lineEdit_file)
        self.btn_choose_file = QtWidgets.QPushButton(parent=self.centralwidget)
        self.btn_choose_file.setObjectName("btn_choose_file")
        self.layoutFile.addWidget(self.btn_choose_file)
        self.verticalLayout.addLayout(self.layoutFile)
        self.layoutFolder = QtWidgets.QHBoxLayout()
        self.layoutFolder.setObjectName("layoutFolder")
        self.lineEdit_folder = QtWidgets.QLineEdit(parent=self.centralwidget)
        self.lineEdit_folder.setObjectName("lineEdit_folder")
        self.layoutFolder.addWidget(self.lineEdit_folder)
        self.btn_choose_folder = QtWidgets.QPushButton(parent=self.centralwidget)
        self.btn_choose_folder.setObjectName("btn_choose_folder")
        self.layoutFolder.addWidget(self.btn_choose_folder)
        self.verticalLayout.addLayout(self.layoutFolder)
        self.lineEdit_filename = QtWidgets.QLineEdit(parent=self.centralwidget)
        self.lineEdit_filename.setObjectName("lineEdit_filename")
        self.verticalLayout.addWidget(self.lineEdit_filename)
        self.label_status = QtWidgets.QLabel(parent=self.centralwidget)
        self.label_status.setObjectName("label_status")
        self.verticalLayout.addWidget(self.label_status)
        self.progressBar = QtWidgets.QProgressBar(parent=self.centralwidget, textVisible=False)
        self.progressBar.setMinimum(0)
        self.progressBar.setMaximum(100)
        self.progressBar.setProperty("value", 0)
        self.progressBar.setObjectName("progressBar")
        self.verticalLayout.addWidget(self.progressBar)
        self.btn_start = QtWidgets.QPushButton(parent=self.centralwidget)
        self.btn_start.setObjectName("btn_start")
        self.verticalLayout.addWidget(self.btn_start)
        MainWindow.setCentralWidget(self.centralwidget)
        self.statusbar = QtWidgets.QStatusBar(parent=MainWindow)
        self.statusbar.setObjectName("statusbar")
        MainWindow.setWindowTitle("ScrapyPrices")
        MainWindow.setStatusBar(self.statusbar)

        self.retranslateUi(MainWindow)
        QtCore.QMetaObject.connectSlotsByName(MainWindow)

    def retranslateUi(self, MainWindow):
        _translate = QtCore.QCoreApplication.translate
        self.btn_choose_file.setText(_translate("MainWindow", "Select file"))
        self.btn_choose_folder.setText(_translate("MainWindow", "Select a folder"))
        self.lineEdit_file.setPlaceholderText(_translate("MainWindow", "File path"))
        self.lineEdit_folder.setPlaceholderText(_translate("MainWindow", "Where to save?"))
        self.lineEdit_filename.setPlaceholderText(_translate("MainWindow", "Enter result file name"))
        self.label_status.setText(_translate("MainWindow", "waiting for process start"))
        self.btn_start.setText(_translate("MainWindow", "Start the process"))

# ==== For UI updates ====
class WorkerSignals(QObject):
    progress = pyqtSignal(int)  
    status = pyqtSignal(str)

class Worker(QRunnable):
    def __init__(self, ui, file_path, file_new_path, file_name):
        super().__init__()
        self.ui = ui
        self.file_path = file_path
        self.file_new_path = file_new_path
        self.file_name = file_name
        self.signals = WorkerSignals()

    def run(self):
        
        # === Disable button ===
        self.ui.btn_start.setEnabled(False)
        
        # === Connect to URL ===
        def connect(url):
            t = 1 # stops in seconds for reconnection
            products = []
            while products == []:
                try:
                    driver.get(url)
                    try:
                        status_search = WebDriverWait(driver, t).until(
                            EC.presence_of_element_located((By.CLASS_NAME, "search-nothing__search-text"))
                        ).text
                        if status_search:# returns an empty list if the product is not found
                            set_status(f"Nothing found for the query [{set_status}] :(")
                            return []
                    except:# if no search status is found, it will try to find product cards
                        products = WebDriverWait(driver, t).until(
                            EC.presence_of_all_elements_located((By.CLASS_NAME, "goods-tile__content"))
                        )
                except Exception as error:
                    print(error)
                    t += 1
                    if t == 15: # if there are internet issues, the program will exit
                        showerror(title="Error", message=f"Connection problems \n [ERROR] {error}")# show error in a popup window
                        exit()
                    set_status(f"reconnecting...")
                    continue
            return products
        
        # === Output status text in the UI ===
        def set_status(text):
            print(f"[] {text}")
            self.signals.status.emit(text)
        
        set_status("Starting Driver...")
        chrome_options = webdriver.ChromeOptions()

        # === Add arguments for stealth mode ===
        # chrome_options.add_argument("--headless")  # Headless mode (no GUI)
        chrome_options.add_argument("--enable-unsafe-swiftshader")
        chrome_options.add_argument("--disable-blink-features=AutomationControlled")  # Hide Selenium (important for bypassing Cloudflare)
        chrome_options.add_argument("--window-size=5,5")  # Window size
        chrome_options.add_argument("--disable-gpu")  # Disable GPU (helps sometimes)
        chrome_options.add_argument("--no-sandbox")  # Useful for Linux environments
        chrome_options.add_argument("--disable-dev-shm-usage")  # For stability in Docker
        
        #  Disable image loading
        chrome_options.add_argument("--blink-settings=imagesEnabled=false")

        #  Disable CSS and fonts loading (unofficial method)
        chrome_options.add_argument("--disable-remote-fonts")
        chrome_options.add_argument("--disable-loading-fonts")
        chrome_options.add_argument("--disable-css")

        # === Launch Chrome with options ===
        driver = webdriver.Chrome(options=chrome_options)

        # === Bypass Cloudflare on Rozetka ===
        # === Remove all cdc_ variables used by Cloudflare bot detection ===
        driver.execute_cdp_cmd("Page.addScriptToEvaluateOnNewDocument", {
            'source': '''
                delete window.cdc_adoQpoasnfa76pfcZLmcfl_Array;
                delete window.cdc_adoQpoasnfa76pfcZLmcfl_JSON;
                delete window.cdc_adoQpoasnfa76pfcZLmcfl_Object;
                delete window.cdc_adoQpoasnfa76pfcZLmcfl_Promise;
                delete window.cdc_adoQpoasnfa76pfcZLmcfl_Proxy;
                delete window.cdc_adoQpoasnfa76pfcZLmcfl_Symbol;
                delete window.cdc_adoQpoasnfa76pfcZLmcfl_Window;
            '''
        })
        
        set_status("Opening an input file...")
        file_type = self.file_path.split(".")[-1]
        try:
            if file_type == "csv": # csv reading
                file_df = pd.read_csv(
                    self.file_path, # file path
                    usecols=[0, 1, 2], # only uses the first 3 columns
                    sep=";", # separator
                    encoding="windows-1251") # encoding
                
            elif file_type == "xlsx": # Excel reading
                file_df = pd.read_excel(
                    self.file_path,
                )
        except Exception as error:
            showerror(title="Error", message=f"Error opening file! \n[ERROR]{error}")
        
        # === Extract product data and store in lists ===
        cids, codes, articals  = (file_df[file_df.keys()[0]],# category ID
                        file_df[file_df.keys()[1]], # product codes 
                        file_df[file_df.keys()[2]]) # articles
        
        data = []
        count = len(codes)
        
        set_status("Starting the Data Collection Process...")
        for i in range(count):
            cid, code, artical = (cids[i],
                codes[i],
                articals[i])
            
            # === Create product URL ===
            url = f'https://rozetka.com.ua/search/?section_id={cid}&sort=cheap&text={artical}'
        
            products = connect(url=url)
            prices = []
            
            if connect:
                # === Loop through each product tile ===
                for product in products:
                    # === Find all elements containing price and availability ===
                    availability = product.find_element(By.CLASS_NAME, "goods-tile__availability").text
                    
                    # === Check if the product is available ===
                    if availability == "Есть в наличии" or availability == "Є в наявності":
                        # === Find the price element ===
                        price = product.find_element(By.CLASS_NAME, "goods-tile__price-value").text
                        prices += [int("".join(price[:-1].split(" ")))] # add to list without currency and convert to integer
                    
            set_status(f"prices found: {len(prices)}")
            
            # === If price list is empty, use NA ===
            if not prices: # no prices
                prices.append("=NA()")
            
            # === Keep only the first 10 prices ===
            if len(prices) >= 10:
                prices = prices[:10]
            else:
                # === Fill remaining slots with empty strings ===
                for i in range(0, 10 - len(prices)):
                    prices += [""]

            # === Create product dictionary === 
            dict_product = {
                "ID categories": cid,
                "Product code": code,
                "Article": artical,
                "URL": url,
            }
            
            # === Add prices to product dictionary ===
            for index, price in enumerate(prices):
                dict_product[f'Price {index+1}'] = price
            
            data += [dict_product]
            set_status(f'Data Collection Processes... [product {code} added]')
            self.signals.progress.emit(int(len(data) / (count / 100)))
        
        set_status("Ending the driver work...")
        driver.quit()
        set_status(f"Collected {len(data)}")
        set_status("Writing data to Excel...")
        df = pd.DataFrame(data)
        file_name = self.file_name if self.file_name.endswith(".xlsx") else self.file_name + ".xlsx"
        df.to_excel(self.file_new_path + file_name, index=False)
        
        set_status("It's done!")
        self.ui.label_status.setStyleSheet("color: #5bed0d;")
        self.signals.progress.emit(100)
        
        # === Re-enable button ===
        self.ui.btn_start.setEnabled(True)

def main():
    
    # === Load CSS file for styles ===
    def load_stylesheet(file_path):
        with open(file_path, "r") as file:
            return file.read()
    
    # === Launch UI ===
    app = QtWidgets.QApplication(sys.argv)
    MainWindow = QtWidgets.QMainWindow()
    ui = Ui_MainWindow()
    ui.setupUi(MainWindow)
    app.setWindowIcon(QtGui.QIcon('icon.ico'))
    app.setStyleSheet(load_stylesheet("css/style.css")) # Load stylesheet
    MainWindow.show()
    
    # === Get input file path ===
    def get_pth_file():
        file_path = fd.askopenfilename()
        ui.lineEdit_file.setText(file_path)
    
    # === Get folder path ===
    def get_pth_folder():
        folder_path = fd.askdirectory()
        ui.lineEdit_folder.setText(folder_path+"/")
    

    # === Start scraping process ===
    def start():
        # === Reset UI effects ===
        ui.lineEdit_file.setStyleSheet("")
        ui.lineEdit_folder.setStyleSheet("")
        ui.lineEdit_filename.setStyleSheet("")
        ui.label_status.setStyleSheet("")
        app.setStyleSheet(load_stylesheet("css/style.css"))
        ui.progressBar.setProperty("value", 0)
        
        file_path = ui.lineEdit_file.text()
        file_new_path = ui.lineEdit_folder.text()
        file_name = ui.lineEdit_filename.text()
        style = "QLineEdit {border-color:#f64848;}"
        
        # === Validate user inputs ===
        if not file_path or not file_new_path or not file_name:
            
            ui.label_status.setStyleSheet("color: #f64848;")
            ui.label_status.setText("All fields must be filled in!")
            
            if not file_path:
                ui.lineEdit_file.setStyleSheet(style)
                
            if not file_new_path:
                ui.lineEdit_folder.setStyleSheet(style)
                
            if not file_name:
                ui.lineEdit_filename.setStyleSheet(style)
            return
        
        # === Check file format compatibility ===
        types = (".xlsx", ".csv")
        if not file_path.endswith(types):
            print(type(file_path))
            ui.label_status.setText("The file does not conform to the read format")
            ui.lineEdit_file.setStyleSheet(style)
            return
        
        worker = Worker(ui, file_path, file_new_path, file_name)
        worker.signals.progress.connect(ui.progressBar.setValue)
        worker.signals.status.connect(ui.label_status.setText)

        QThreadPool.globalInstance().start(worker)
        
    # === Bind buttons ===
    ui.btn_choose_file.clicked.connect(get_pth_file)
    ui.btn_choose_folder.clicked.connect(get_pth_folder)
    ui.btn_start.clicked.connect(start)
    
    # === Exit application on window close ===
    sys.exit(app.exec())

# === Check if script is run as main ===
if __name__ == '__main__':
    main()
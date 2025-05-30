import sys
import os
import re
import time
import math
import pandas as pd
import csv
import copy
import numpy as np
import win32com.client
import datetime
# from PyQt5 import QtCore, QtGui, QtWidgets
# from PyQt5.QtWidgets import QApplication, QMainWindow
from PyQt5.QtWidgets import QApplication, QFileDialog, QMainWindow, QMessageBox, QVBoxLayout, QPushButton, QAction
from PyQt5.QtCore import QDate
from PyQt5.QtGui import QIcon
from Get_Data import *
from File_Operate import *
from Sap_Function import *
from Hour_Operate_Ui import Ui_MainWindow
from Data_Table import *
from Logger import *
from theme_manager_theme import ThemeManager





class MyMainWindow(QMainWindow, Ui_MainWindow):
    def __init__(self, parent=None):
        super(MyMainWindow, self).__init__(parent)
        self.setupUi(self)

        self.theme_manager = ThemeManager(QApplication.instance())
        self.init_theme_action()

        self.setGeometry(100, 100, 300, 200)

        self.theme_manager = ThemeManager(QApplication.instance())

        layout = QVBoxLayout()

        toggle_button = QPushButton("Toggle Theme")
        toggle_button.clicked.connect(self.theme_manager.toggle_theme)
        layout.addWidget(toggle_button)

        self.setMinimumSize(1200, 750)

        self.actionExport.triggered.connect(self.exportConfig)
        self.actionImport.triggered.connect(self.importConfig)
        self.actionExit.triggered.connect(MyMainWindow.close)
        self.actionHelp.triggered.connect(self.showVersion)
        self.actionAuthor.triggered.connect(self.showAuthorMessage)
        self.theme_manager.set_theme("blue")  # 设置默认主题
        self.pushButton_78.clicked.connect(lambda: self.getHourFile(configContent['Hour_Files_Import_URL']))
        self.pushButton_79.clicked.connect(self.hourOperate)
        self.pushButton.clicked.connect(lambda: self.viewOdmData(self.lineEdit_31.text()))
        self.filesUrl = []

    def init_theme_action(self):
        theme_action = QAction(QIcon('theme_icon.png'), 'Toggle Theme', self)
        theme_action.setStatusTip('Toggle Theme')
        theme_action.triggered.connect(self.toggle_theme)

        # 将 action 添加到菜单（如果有的话）
        if hasattr(self, 'menuBar'):
            view_menu = self.menuBar().addMenu('Theme')
            view_menu.addAction(theme_action)

        # # 将 action 添加到工具栏
        # toolbar = self.addToolBar('主题')
        # toolbar.addAction(theme_action)

    def toggle_theme(self):
        self.theme_manager.set_random_theme()
        # 可以在这里添加其他需要在主题切换后更新的UI元素

    def getConfig(self):
        # 初始化，获取或生成配置文件
        global configFileUrl
        global desktopUrl
        global now
        global last_time
        global today
        global oneWeekday
        global fileUrl

        date = datetime.datetime.now() + datetime.timedelta(days=1)
        now = int(time.strftime('%Y'))
        last_time = now - 1
        today = time.strftime('%Y.%m.%d')
        oneWeekday = (datetime.datetime.now() + datetime.timedelta(days=7)).strftime('%Y.%m.%d')
        desktopUrl = os.path.join(os.path.expanduser("~"), 'Desktop')
        configFileUrl = '%s/config' % desktopUrl
        configFile = os.path.exists('%s/config_hour.csv' % configFileUrl)
        # print(desktopUrl,configFileUrl,configFile)
        if not configFile:  # 判断是否存在文件夹如果不存在则创建为文件夹
            reply = QMessageBox.question(self, '信息', '确认是否要创建配置文件', QMessageBox.Yes | QMessageBox.No, QMessageBox.Yes)
            if reply == QMessageBox.Yes:
                if not os.path.exists(configFileUrl):
                    os.makedirs(configFileUrl)
                MyMainWindow.createConfigContent(self)
                MyMainWindow.getConfigContent(self)
                self.textBrowser.append("创建并导入配置成功")
            else:
                exit()
        else:
            MyMainWindow.getConfigContent(self)

    # 获取配置文件内容
    def getConfigContent(self):
        # 配置文件
        csvFile = pd.read_csv('%s/config_hour.csv' % configFileUrl, names=['A', 'B', 'C'])
        global configContent
        global username
        global role
        global staff_dict
        configContent = {}
        staff_dict = {}
        # configContent[configContent.get('Business_Department','CS')] = []
        # configContent[configContent.get('Lab_1','PHY')] = []
        # configContent[configContent.get('Lab_2','CHM')] = []
        username = list(csvFile['A'])
        number = list(csvFile['B'])
        role = list(csvFile['C'])
        for i in range(len(username)):
            configContent['%s' % username[i]] = number[i]
            if role[i] == configContent.get('Business_Department', 'CS'):
                # 使用 setdefault 确保键存在且为列表类型
                staff_dict.setdefault(configContent.get('Business_Department', 'CS'), []).append(username[i])
            if role[i] == configContent.get('Lab_1', 'PHY'):
                # 使用 setdefault 确保键存在且为列表类型
                staff_dict.setdefault(configContent.get('Lab_1', 'PHY'), []).append(username[i])
            if role[i] == configContent.get('Lab_2', 'CHM'):
                # 使用 setdefault 确保键存在且为列表类型
                staff_dict.setdefault(configContent.get('Lab_2', 'CHM'), []).append(username[i])

        try:
            self.textBrowser_4.append("配置获取成功")
        except AttributeError:
            QMessageBox.information(self, "提示信息", "已获取配置文件内容", QMessageBox.Yes)
        else:
            pass

    # 创建配置文件
    def createConfigContent(self):
        global monthAbbrev
        months = "JanFebMarAprMayJunJulAugSepOctNovDec"
        n = time.strftime('%m')
        pos = (int(n) - 1) * 3
        monthAbbrev = months[pos:pos + 3]

        configContent = [
            ['Hour_Files_Import_URL', "N:\\XM Softlines\\6. Personel\\5. Personal\\Supporting Team\\2.财务\\2.SAP\\1.ODM Data - XM\\3.Hours",'Invoice文件导入路径'],
            ['Hour_Files_Export_URL', "N:\\XM Softlines\\6. Personel\\5. Personal\\Supporting Team\\2.财务\\2.SAP\\1.ODM Data - XM\\3.Hours",'Invoice文件导入路径'],
            ['Hour_Field_Mapping', "{'staff_id': 'staff_id','week': 'week','order_no': 'order_no','allocated_hours': 'allocated_hours','office_time':'office_time','material_code': 'material_code','item': 'item','allocated_day': 'allocated_day','staff_name': 'staff_name'}",'对应字段映射'],
        ]
        config = np.array(configContent)
        df = pd.DataFrame(config)
        df.to_csv('%s/config_hour.csv' % configFileUrl, index=0, header=0, encoding='utf_8_sig')
        self.textBrowser_4.append("配置文件创建成功")
        QMessageBox.information(self, "提示信息",
                                "默认配置文件已经创建好，\n如需修改请在用户桌面查找config文件夹中config_hour.csv，\n将相应的文件内容替换成用户需求即可，修改后记得重新导入配置文件。",
                                QMessageBox.Yes)

    # 导出配置文件
    def exportConfig(self):
        # 重新导出默认配置文件
        reply = QMessageBox.question(self, '信息', '确认是否要创建默认配置文件', QMessageBox.Yes | QMessageBox.No, QMessageBox.Yes)
        if reply == QMessageBox.Yes:
            MyMainWindow.createConfigContent(self)
        else:
            QMessageBox.information(self, "提示信息", "没有创建默认配置文件，保留原有的配置文件", QMessageBox.Yes)

    # 导入配置文件
    def importConfig(self):
        # 重新导入配置文件
        reply = QMessageBox.question(self, '信息', '确认是否要导入配置文件', QMessageBox.Yes | QMessageBox.No, QMessageBox.Yes)
        if reply == QMessageBox.Yes:
            MyMainWindow.getConfigContent(self)
        else:
            QMessageBox.information(self, "提示信息", "没有重新导入配置文件，将按照原有的配置文件操作", QMessageBox.Yes)

    # 界面设置默认配置文件信息

    def showAuthorMessage(self):
        # 关于作者
        QMessageBox.about(self, "关于",
                          "人生苦短，码上行乐。\n\n\n        ----Frank Chen")

    def showVersion(self):
        # 关于作者
        QMessageBox.about(self, "版本",
                          "V 22.01.11\n\n\n 2022-04-26")

    def getAmountVat(self):
        amount = float(self.doubleSpinBox_2.text())
        self.doubleSpinBox_4.setValue(amount * 1.06)

    # 获取文件
    def getFile(self, path):
        selectBatchFile = QFileDialog.getOpenFileName(self, '选择ODM导出文件',
                                                      '%s\\%s' % (path, today),
                                                      'files(*.docx;*.xls*;*.csv)')
        fileUrl = selectBatchFile[0]
        return fileUrl

    def getHourFile(self, path):
        fileUrl = self.getFile(path)
        if fileUrl:
            self.lineEdit_31.setText(fileUrl)
            self.textBrowser_4.append("文件路径获取成功")
            return fileUrl
        else:
            self.textBrowser_4.append("未选择文件")
            return None

    # 查看SAP操作数据详情
    def viewOdmData(self, path):
        fileUrl = path
        odm_data_obj = Get_Data()
        df = odm_data_obj.getFileTableData(fileUrl)
        myTable.createTable(df)
        myTable.showMaximized()

    def hourOperate(self):
        """
        处理工时数据并进行SAP操作
        """

        time_str = datetime.datetime.now().strftime("%Y%m%d_%H%M%S")
        log_file = os.path.join(configContent['Hour_Files_Export_URL'], f'log_{time_str}.xlsx')
        columns = [
            'timestamp',
            'staff_id',
            'week',
            'order_no',
            'allocated_hours',
            'office_time',
            'material_code',
            'item',
            'allocated_day',
            'staff_name',
            'status',
            'message',
            'Update'
        ]
        logger = Logger(log_file=log_file, columns=columns)
        try:
            # 获取文件路径
            hour_path = self.lineEdit_31.text()
            if not hour_path:
                QMessageBox.warning(self, "警告", "请先选择工时文件！")
                return

            log_data = []  # 用于存储日志数据
            
            # 获取并处理数据
            get_data = Get_Data()
            raw_data = get_data.getFileTableData(hour_path)
            
            # 重命名字段
            renamed_data = get_data.rename_hour_fields(raw_data, configContent['Hour_Field_Mapping'])
            
            # # 按staff_id和week分组
            # grouped_data = get_data.group_hour_data(renamed_data)
            
            # 初始化SAP操作对象
            sap = Sap()
            
            # 记录当前处理的staff_id和week
            current_staff_id = None
            current_week = None
            is_first_login = True  # 标记是否是第一次登录
            
            # 遍历分组后的数据
            for _, row in renamed_data.iterrows():
                staff_id = row['staff_id']
                week = row['week']
                
                # 如果staff_id或week发生变化，需要重新登录
                if staff_id != current_staff_id or week != current_week:
                    # 如果不是第一次登录，需要先保存之前的工时
                    if not is_first_login:
                        if not sap.save_hours():
                            error_msg = f"保存工时失败！Staff ID: {current_staff_id}, Week: {current_week}"
                            # logger.error(error_msg)
                            log_data.append({
                                'timestamp': datetime.datetime.now(),
                                'staff_id': current_staff_id,
                                'week': current_week,
                                'status': 'Failed',
                                'message': error_msg
                            })
                            continue
                    
                    # 登录SAP
                    if not sap.login_hour_gui(row):
                        error_msg = f"登录SAP失败！Staff ID: {staff_id}, Week: {week}"
                        # logger.error(error_msg)
                        log_data.append({
                            'timestamp': datetime.datetime.now(),
                            'staff_id': staff_id,
                            'week': week,
                            'status': 'Failed',
                            'message': error_msg
                        })
                        continue
                    
                    current_staff_id = staff_id
                    current_week = week
                    is_first_login = False
                
                # 记录工时
                try:
                    # 准备工时数据
                    # hour_data = {
                    #     'staff_id': staff_id,
                    #     'week': week,
                    #     'order_no': row['order_no'],
                    #     'hours': row['hours'],
                    #     'department': row['department'],
                    #     'project': row['project'],
                    #     'description': row['description']
                    # }
                    hour_data = row
                    # 调用recording_hours方法记录工时
                    if not sap.recording_hours(hour_data):
                        error_msg = f"记录工时失败！Staff ID: {staff_id}, Week: {week}"
                        # logger.error(error_msg)
                        log_data.append({
                            'timestamp': datetime.datetime.now(),
                            'staff_id': staff_id,
                            'week': week,
                            'status': 'Failed',
                            'message': error_msg
                        })
                        continue
                    
                    success_msg = f"成功处理 Staff ID: {staff_id}, Week: {week} 的工时数据"
                    # logger.info(success_msg)
                    log_data.append({
                        'timestamp': datetime.datetime.now(),
                        'staff_id': staff_id,
                        'week': week,
                        'status': 'Success',
                        'message': success_msg
                    })
                    
                    self.textBrowser_4.append(success_msg)
                    
                except Exception as e:
                    error_msg = f"处理工时数据时出错: {str(e)}"
                    # logger.error(error_msg)
                    log_data.append({
                        'timestamp': datetime.datetime.now(),
                        'staff_id': staff_id,
                        'week': week,
                        'status': 'Failed',
                        'message': error_msg
                    })
                    continue
            
            # 最后一次保存
            if not is_first_login:
                if not sap.save_hours():
                    error_msg = f"最后一次保存工时失败！Staff ID: {current_staff_id}, Week: {current_week}"
                    # logger.error(error_msg)
                    log_data.append({
                        'timestamp': datetime.datetime.now(),
                        'staff_id': current_staff_id,
                        'week': current_week,
                        'status': 'Failed',
                        'message': error_msg
                    })
            
            # 将日志数据保存为Excel文件
            log_df = pd.DataFrame(log_data)
            log_file_path = os.path.join(os.path.dirname(hour_path), f'hour_operation_log_{datetime.datetime.now().strftime("%Y%m%d_%H%M%S")}.xlsx')
            log_df.to_excel(log_file_path, index=False)
            
            QMessageBox.information(self, "完成", f"所有工时数据处理完成！\n日志文件保存在：{log_file_path}")
            
        except Exception as e:
            error_msg = f"处理过程中出现错误: {str(e)}"
            # logger.error(error_msg)
            QMessageBox.critical(self, "错误", error_msg)




if __name__ == "__main__":
    QtCore.QCoreApplication.setAttribute(QtCore.Qt.AA_EnableHighDpiScaling)
    app = QApplication(sys.argv)
    myWin = MyMainWindow()
    myTable = MyTableWindow()
    myWin.show()
    myWin.getConfig()
    sys.exit(app.exec_())

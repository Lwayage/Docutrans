# -*- coding: utf-8 -*-
from PyQt5 import QtWidgets
from PyQt5.QtCore import Qt, QCoreApplication
from PyQt5.QtGui import QIcon
from win32com.client import gencache
from win32com.client import constants, gencache
from FilenameSort import sort_list_by_name as FilenameSort
import sys, os
import fitz, glob

class MainUi(QtWidgets.QMainWindow):
    def __init__(self):
        super().__init__()
        self.init_ui()

    def init_ui(self):
        self.m_flag = False                                     # “鼠标可拖动无边框窗口”中，变量m_FLag的初始化
        self.t_flag = 0                                         # “更改转换模式”中，变量t_flag的初始化
        self.setFixedSize(800,520)                              # 设置主窗口大小
        self.main_widget = QtWidgets.QWidget()                  # 创建窗口主部件
        self.main_layout = QtWidgets.QGridLayout()              # 创建窗口主部件网格布局
        self.main_widget.setLayout(self.main_layout)            # 将网格布局应用到窗口主部件上
        self.setCentralWidget(self.main_widget)                 # 设置窗口主部件为首要部件

        self.right_widget = QtWidgets.QWidget()                 # 创建右侧部件
        self.right_layout = QtWidgets.QGridLayout()             # 创建网格布局
        self.right_widget.setLayout(self.right_layout)          # 将网格布局应用到右侧部件上
        # 将 右侧部件 添加到 主部件布局 上， 右侧部件坐标(0, 2)，占1行10列
        self.main_layout.addWidget(self.right_widget, 0, 2, 1, 10)

        self.Menu()                                             # 调用菜单
        self.TopRight()                                         # 调用右上三个按钮
        self.PDF_Page()                                         # 调用PDF
        self.Word_Page()
        self.Help_Page()
        self.ChangePage('PDF')

        self.setWindowTitle('(。・∀・)ノ')                      # 设置窗口标题
        self.setWindowIcon(QIcon('./photos/1.4.ico'))           # 设置窗口图标
        self.setAttribute(Qt.WA_TranslucentBackground)          # 设置窗口背景透明
        self.setWindowFlag(Qt.FramelessWindowHint)              # 隐藏边框
        self.main_layout.setSpacing(0)                          # 去除缝隙

        self.StyleSheet()                                       # 调用样式表单
    
    def Menu(self):
        self.menu_widget = QtWidgets.QWidget()                  # 创建菜单部件
        self.menu_layout = QtWidgets.QVBoxLayout()              # 创建菜单部件垂直布局
        self.menu_widget.setLayout(self.menu_layout)            # 将垂直布局应用到菜单部件上
        # 将菜单部件添加到主窗口部件的布局上， 菜单部件坐标(0, 0)， 占1行2列
        self.main_layout.addWidget(self.menu_widget, 0, 0, 1, 2)
        
        # 为菜单部件创建按钮
        self.menu_button_0 = QtWidgets.QPushButton('文档转换图片工具')
        self.menu_button_1 = QtWidgets.QPushButton()
        self.menu_button_2 = QtWidgets.QPushButton()
        self.menu_button_3 = QtWidgets.QPushButton()

        self.menu_button_1.clicked.connect(lambda:self.ChangePage('PDF')) # 按钮1链接到对应页面
        self.menu_button_2.clicked.connect(lambda:self.ChangePage('Word')) # 按钮1链接到对应页面
        self.menu_button_3.clicked.connect(lambda:self.ChangePage('Help')) # 按钮1链接到对应页面

        self.menu_button_0.setObjectName('label')               # 为按钮设置一个标签
        self.menu_button_1.setObjectName('PDF')
        self.menu_button_2.setObjectName('Word')
        self.menu_button_3.setObjectName('Help')

        self.menu_button_0.setFixedWidth(120)                   # 限制按钮的长度
        self.menu_button_1.setFixedWidth(120)
        # 规定按钮可延伸
        self.menu_button_1.setSizePolicy(QtWidgets.QSizePolicy.Expanding, QtWidgets.QSizePolicy.Expanding)

        self.menu_button_2.setFixedWidth(120)
        self.menu_button_2.setSizePolicy(QtWidgets.QSizePolicy.Expanding, QtWidgets.QSizePolicy.Expanding)

        self.menu_button_3.setFixedWidth(120)
        self.menu_button_3.setSizePolicy(QtWidgets.QSizePolicy.Expanding, QtWidgets.QSizePolicy.Expanding)

        self.menu_layout.addWidget(self.menu_button_0)          # 将创建的按钮添加到菜单的布局上
        self.menu_layout.addWidget(self.menu_button_1)
        self.menu_layout.addWidget(self.menu_button_2)
        self.menu_layout.addWidget(self.menu_button_3)
    
    def TopRight(self):
        self.top_widget = QtWidgets.QWidget()                  # 创建窗口主部件
        self.top_layout = QtWidgets.QGridLayout()              # 创建窗口主部件网格布局
        self.top_widget.setLayout(self.top_layout)            # 将网格布局应用到窗口主部件上
        self.right_layout.addWidget(self.top_widget, 0, 0, 1, 1, Qt.AlignRight | Qt.AlignTop)

        self.close_button = QtWidgets.QPushButton()
        self.visit_button = QtWidgets.QPushButton()
        self.minim_button = QtWidgets.QPushButton()
        self.close_button.setFixedSize(20, 20)
        self.visit_button.setFixedSize(20, 20)
        self.minim_button.setFixedSize(20, 20)
        self.close_button.setObjectName('close')
        self.visit_button.setObjectName('visit')
        self.minim_button.setObjectName('minim')
        self.close_button.clicked.connect(QCoreApplication.instance().quit)
        self.minim_button.clicked.connect(lambda:self.setWindowState(Qt.WindowMinimized))
        self.top_layout.addWidget(self.minim_button, 0, 0, 1, 1)
        self.top_layout.addWidget(self.visit_button, 0, 1, 1, 1)
        self.top_layout.addWidget(self.close_button, 0, 2, 1, 1)

    def PDF_Page(self):
        self.pdf_widget = QtWidgets.QWidget()                   # 创建 PDF部件
        self.pdf_layout = QtWidgets.QGridLayout()               # 创建 网格布局
        self.pdf_widget.setLayout(self.pdf_layout)              # 将 网格布局 应用到 PDF部件 上
        # 将 PDF部件 添加到 右侧布局 上， PDF部件坐标(1, 0)，占10行1列
        self.right_layout.addWidget(self.pdf_widget, 1, 0, 10, 1)

        self.pdf_widget_1 = QtWidgets.QWidget()                 # 创建 PDF部件1
        self.pdf_layout_1 = QtWidgets.QHBoxLayout()             # 创建 水平布局1
        self.pdf_widget_1.setLayout(self.pdf_layout_1)          # 将 水平布局1 应用到 PDF部件1 上
        # 将 PDF部件1 添加到 PDF布局 上
        self.pdf_layout.addWidget(self.pdf_widget_1, 0, 0, 1, 1)

        self.pdf_file_label = QtWidgets.QLabel('选择PDF文件：')
        self.pdf_file_label.setFixedWidth(88)
        self.pdf_layout_1.addWidget(self.pdf_file_label)

        self.pdf_file_lineedit = QtWidgets.QLineEdit()
        self.pdf_layout_1.addWidget(self.pdf_file_lineedit)

        self.pdf_widget_2 = QtWidgets.QWidget()                 # 创建 PDF部件2
        self.pdf_layout_2 = QtWidgets.QHBoxLayout()             # 创建 水平布局2
        self.pdf_widget_2.setLayout(self.pdf_layout_2)          # 将 水平布局2 应用到 PDF部件2 上
        # 将 PDF部件2 添加到 PDF布局 上
        self.pdf_layout.addWidget(self.pdf_widget_2, 1, 0, 1, 1)

        self.pdf_remind_label_1 = QtWidgets.QLabel('请注意文件格式或路径是否正确')
        self.pdf_layout_2.addWidget(self.pdf_remind_label_1)

        self.pdf_button_1 = QtWidgets.QPushButton("浏览")
        self.pdf_button_1.clicked.connect(lambda:self.Openfile('PDF'))
        self.pdf_button_1.setObjectName('button')
        self.pdf_button_1.setFixedSize(110, 26)
        self.pdf_layout_2.addWidget(self.pdf_button_1, 0, Qt.AlignRight)

        self.pdf_widget_3 = QtWidgets.QWidget()                 # 创建 PDF部件3
        self.pdf_layout_3 = QtWidgets.QVBoxLayout()             # 创建 垂直布局3
        self.pdf_widget_3.setLayout(self.pdf_layout_3)          # 将 垂直布局3 应用到 PDF部件3 上
        # 将 PDF部件3 添加到 PDF布局 上
        self.pdf_layout.addWidget(self.pdf_widget_3, 2, 0, 1, 1)

        self.pdf_blank_label = QtWidgets.QLabel('')
        self.pdf_layout_3.addWidget(self.pdf_blank_label)

        self.pdf_widget_4 = QtWidgets.QWidget()                 # 创建 PDF部件4
        self.pdf_layout_4 = QtWidgets.QHBoxLayout()             # 创建 水平布局4
        self.pdf_widget_4.setLayout(self.pdf_layout_4)          # 将 水平布局4 应用到 PDF部件4 上
        # 将 PDF部件4 添加到 PDF布局 上
        self.pdf_layout.addWidget(self.pdf_widget_4, 3, 0, 1, 1)

        self.pdf_save_label = QtWidgets.QLabel('图片保存位置：')
        self.pdf_save_label.setFixedWidth(88)
        self.pdf_layout_4.addWidget(self.pdf_save_label)

        self.pdf_save_lineedit = QtWidgets.QLineEdit()
        self.pdf_layout_4.addWidget(self.pdf_save_lineedit)

        self.pdf_widget_5 = QtWidgets.QWidget()                 # 创建 PDF部件5
        self.pdf_layout_5 = QtWidgets.QHBoxLayout()             # 创建 水平布局5
        self.pdf_widget_5.setLayout(self.pdf_layout_5)          # 将 水平布局5 应用到 PDF部件5 上
        # 将 PDF部件5 添加到 PDF布局 上
        self.pdf_layout.addWidget(self.pdf_widget_5, 4, 0, 1, 1)

        self.pdf_remind_label_2 = QtWidgets.QLabel('默认保存位置与原文档相同')
        self.pdf_layout_5.addWidget(self.pdf_remind_label_2)

        self.pdf_button_2 = QtWidgets.QPushButton("浏览")
        self.pdf_button_2.clicked.connect(lambda:self.OpenPath('PDF'))
        self.pdf_button_2.setObjectName('button')
        self.pdf_button_2.setFixedSize(110, 26)
        self.pdf_layout_5.addWidget(self.pdf_button_2, 0, Qt.AlignRight)

        self.pdf_widget_6 = QtWidgets.QWidget()                 # 创建 PDF部件6
        self.pdf_layout_6 = QtWidgets.QGridLayout()             # 创建 网格布局6
        self.pdf_widget_6.setLayout(self.pdf_layout_6)          # 将 网格布局6 应用到 PDF部件6 上
        # 将 PDF部件6 添加到 PDF布局 上
        self.pdf_layout.addWidget(self.pdf_widget_6, 5, 0, 3, 1)

        self.pdf_mode_label = QtWidgets.QLabel('转换模式:')
        self.pdf_mode_label.setAlignment(Qt.AlignRight | Qt.AlignCenter)
        self.pdf_layout_6.addWidget(self.pdf_mode_label, 0, 0, 1, 1)

        self.pdf_combobox = QtWidgets.QComboBox()
        self.pdf_combobox.currentIndexChanged.connect(self.ChangeMode)
        self.pdf_combobox.setFixedSize(100, 20)
        self.pdf_combobox.addItems(['PDF转图片', '图片转PDF'])
        self.pdf_layout_6.addWidget(self.pdf_combobox, 0, 1, 1, 1)

        self.pdf_radiobutton = QtWidgets.QRadioButton('转换完成后打开文件夹')
        self.pdf_radiobutton.setFixedSize(150, 20)
        self.pdf_layout_6.addWidget(self.pdf_radiobutton, 0, 2, 1, 1, Qt.AlignCenter)

        self.pdf_button_3 = QtWidgets.QPushButton("转换")
        self.pdf_button_3.clicked.connect(self.Transform)
        self.pdf_button_3.setObjectName('button')
        self.pdf_button_3.setFixedSize(110, 26)
        self.pdf_layout_6.addWidget(self.pdf_button_3, 0, 3, 1, 1)

        self.pdf_widget_7 = QtWidgets.QWidget()                 # 创建 PDF部件7
        self.pdf_layout_7 = QtWidgets.QHBoxLayout()             # 创建 水平布局7
        self.pdf_widget_7.setLayout(self.pdf_layout_7)          # 将 水平布局7 应用到 PDF部件7 上
        # 将 PDF部件7 添加到 PDF布局 上
        self.pdf_layout.addWidget(self.pdf_widget_7, 9, 0, 3, 1)

        self.pdf_progressbar = QtWidgets.QProgressBar()
        self.pdf_progressbar.setAlignment(Qt.AlignCenter)
        self.pdf_progressbar.setFixedHeight(21)
        self.pdf_progressbar.setVisible(False)
        self.pdf_layout_7.addWidget(self.pdf_progressbar)

    def Word_Page(self):
        self.word_widget = QtWidgets.QWidget()                   # 创建 Word部件
        self.word_layout = QtWidgets.QGridLayout()               # 创建 网格布局
        self.word_widget.setLayout(self.word_layout)              # 将 网格布局 应用到 Word部件 上
        # 将 Word部件 添加到 右侧布局 上， Word部件坐标(1, 0)，占10行1列
        self.right_layout.addWidget(self.word_widget, 1, 0, 10, 1)

        self.word_widget_1 = QtWidgets.QWidget()                 # 创建 Word部件1
        self.word_layout_1 = QtWidgets.QHBoxLayout()             # 创建 水平布局1
        self.word_widget_1.setLayout(self.word_layout_1)          # 将 水平布局1 应用到 Word部件1 上
        # 将 Word部件1 添加到 Word布局 上
        self.word_layout.addWidget(self.word_widget_1, 0, 0, 1, 1)

        self.word_file_label = QtWidgets.QLabel('选择Word文件：')
        self.word_file_label.setFixedWidth(88)
        self.word_layout_1.addWidget(self.word_file_label)

        self.word_file_lineedit = QtWidgets.QLineEdit()
        self.word_layout_1.addWidget(self.word_file_lineedit)

        self.word_widget_2 = QtWidgets.QWidget()                 # 创建 Word部件2
        self.word_layout_2 = QtWidgets.QHBoxLayout()             # 创建 水平布局2
        self.word_widget_2.setLayout(self.word_layout_2)          # 将 水平布局2 应用到 Word部件2 上
        # 将 Word部件2 添加到 Word布局 上
        self.word_layout.addWidget(self.word_widget_2, 1, 0, 1, 1)

        self.word_remind_label_1 = QtWidgets.QLabel('请注意文件格式或路径是否正确')
        self.word_layout_2.addWidget(self.word_remind_label_1)

        self.word_button_1 = QtWidgets.QPushButton("浏览")
        self.word_button_1.clicked.connect(lambda:self.Openfile('Word'))
        self.word_button_1.setObjectName('button')
        self.word_button_1.setFixedSize(110, 26)
        self.word_layout_2.addWidget(self.word_button_1, 0, Qt.AlignRight)

        self.word_widget_3 = QtWidgets.QWidget()                 # 创建 Word部件3
        self.word_layout_3 = QtWidgets.QVBoxLayout()             # 创建 垂直布局3
        self.word_widget_3.setLayout(self.word_layout_3)          # 将 垂直布局3 应用到 Word部件3 上
        # 将 Word部件3 添加到 Word布局 上
        self.word_layout.addWidget(self.word_widget_3, 2, 0, 1, 1)

        self.word_blank_label = QtWidgets.QLabel('')
        self.word_layout_3.addWidget(self.word_blank_label)

        self.word_widget_4 = QtWidgets.QWidget()                 # 创建 Word部件4
        self.word_layout_4 = QtWidgets.QHBoxLayout()             # 创建 水平布局4
        self.word_widget_4.setLayout(self.word_layout_4)          # 将 水平布局4 应用到 Word部件4 上
        # 将 Word部件4 添加到 Word布局 上
        self.word_layout.addWidget(self.word_widget_4, 3, 0, 1, 1)

        self.word_save_label = QtWidgets.QLabel('图片保存位置：')
        self.word_save_label.setFixedWidth(88)
        self.word_layout_4.addWidget(self.word_save_label)

        self.word_save_lineedit = QtWidgets.QLineEdit()
        self.word_layout_4.addWidget(self.word_save_lineedit)

        self.word_widget_5 = QtWidgets.QWidget()                 # 创建 Word部件5
        self.word_layout_5 = QtWidgets.QHBoxLayout()             # 创建 水平布局5
        self.word_widget_5.setLayout(self.word_layout_5)          # 将 水平布局5 应用到 Word部件5 上
        # 将 Word部件5 添加到 Word布局 上
        self.word_layout.addWidget(self.word_widget_5, 4, 0, 1, 1)

        self.word_remind_label_2 = QtWidgets.QLabel('默认保存位置与原文档相同')
        self.word_layout_5.addWidget(self.word_remind_label_2)

        self.word_button_2 = QtWidgets.QPushButton("浏览")
        self.word_button_2.clicked.connect(lambda:self.OpenPath('Word'))
        self.word_button_2.setObjectName('button')
        self.word_button_2.setFixedSize(110, 26)
        self.word_layout_5.addWidget(self.word_button_2, 0, Qt.AlignRight)

        self.word_widget_6 = QtWidgets.QWidget()                 # 创建 Word部件6
        self.word_layout_6 = QtWidgets.QGridLayout()             # 创建 网格布局6
        self.word_widget_6.setLayout(self.word_layout_6)          # 将 网格布局6 应用到 Word部件6 上
        # 将 Word部件6 添加到 Word布局 上
        self.word_layout.addWidget(self.word_widget_6, 5, 0, 3, 1)

        self.word_radiobutton = QtWidgets.QRadioButton('转换完成后打开文件夹')
        self.word_radiobutton.setFixedSize(150, 20)
        self.word_layout_6.addWidget(self.word_radiobutton, 0, 2, 1, 1, Qt.AlignRight)

        self.word_button_3 = QtWidgets.QPushButton("转换")
        self.word_button_3.clicked.connect(self.W_Transform)
        self.word_button_3.setObjectName('button')
        self.word_button_3.setFixedSize(110, 26)
        self.word_layout_6.addWidget(self.word_button_3, 0, 3, 1, 1)

        self.word_widget_7 = QtWidgets.QWidget()                 # 创建 Word部件7
        self.word_layout_7 = QtWidgets.QHBoxLayout()             # 创建 水平布局7
        self.word_widget_7.setLayout(self.word_layout_7)          # 将 水平布局7 应用到 Word部件7 上
        # 将 Word部件7 添加到 Word布局 上
        self.word_layout.addWidget(self.word_widget_7, 9, 0, 3, 1)

        self.word_progressbar = QtWidgets.QProgressBar()
        self.word_progressbar.setAlignment(Qt.AlignCenter)
        self.word_progressbar.setFixedHeight(21)
        self.word_progressbar.setVisible(False)
        self.word_layout_7.addWidget(self.word_progressbar)

    def Help_Page(self):
        self.help_widget = QtWidgets.QWidget()
        self.help_layout = QtWidgets.QGridLayout()
        self.help_widget.setLayout(self.help_layout)
        self.help_test = QtWidgets.QLabel(".. ..-. -.-- --- ..- .... .- ...- . .- -. -.-- .--. .-. --- -... .-.. . -- ... - .-. -.-- - --- -.-. --- -. - .- -.-. - -- . --- -. --.- --.- .---- --... ...-- .---- .---- ..--- ....- ..... ....- --...")
        self.help_test.setAlignment(Qt.AlignRight)
        self.help_test.setTextInteractionFlags(Qt.TextSelectableByMouse)
        self.help_layout.addWidget(self.help_test, 0, 0, Qt.AlignRight | Qt.AlignBottom)
        self.right_layout.addWidget(self.help_widget, 1, 0, 10, 1)

    def ChangePage(self, pagename):
        pagedir = {'PDF': 0, 'Word': 0, 'Help': 0}
        pagedir[pagename] = 1
        self.pdf_widget.setVisible(pagedir['PDF'])
        self.word_widget.setVisible(pagedir['Word'])
        self.help_widget.setVisible(pagedir['Help'])

    def StyleSheet(self):
        # 为右上三个按钮设定样式
        self.top_widget.setStyleSheet('''
            QPushButton#close{
                background:#F76677;
                border-radius:10px;
            }
            QPushButton#close:hover{background:#f34257;}

            QPushButton#visit{
                background:#F7D674;
                border-radius:10px;
            }
            QPushButton#visit:hover{background:#f6ff61;}
            
            QPushButton#minim{
                background:#6DDF6D;
                border-radius:10px;
            }
            QPushButton#minim:hover{background:#74ff81;}
        ''')

        # 为菜单各个部件设定样式
        self.menu_widget.setStyleSheet('''
            QWidget{background:#282828;}
            QPushButton{border:none; color:white;}
            QPushButton#label{
                border:none;
                border-bottom:1px solid white;
                font-size:11px;
                font-weight:700;
                font-family: Microsoft YaHei;
            }
            QPushButton#PDF{image: url(./photos/PDF.png);}
            QPushButton#Word{image: url(./photos/DOC.png);}
            QPushButton#Help{image: url(./photos/HELP.png);}
            QPushButton#PDF:hover{border-left:4px solid #ff3259;}
            QPushButton#Word:hover{border-left:4px solid #2060f0;}
            QPushButton#Help:hover{border-left:4px solid #efefef;}
        ''')

        # 为右边页面各个部件设定样式
        self.right_widget.setStyleSheet('''
            QWidget{background:#535353;}
            QLabel{
                color: white;
                font-size:13px;
                font-family:Microsoft YaHei;
            }
            QPushButton#button{
                border-radius:13px;
                border:2px solid #bebebe;
                color: white;
                font-family: Microsoft YaHei;
                background:#535353;
            }
            QPushButton#button:hover{
                border-radius:13px;
                border:2px solid #3defff;
                color: white;
                font-family: Microsoft YaHei;
                background:#535353;
            }
            QPushButton#button:pressed{
                border:none;
                border-radius:13px;
                color: black;
                font-family: Microsoft YaHei;
                background:white;
            }
            QLineEdit{
                border:1px solid #999999;
                color:white;
                font-family: Microsoft YaHei;
                background:#454545;
            }
            QComboBox{
                border:1px solid #666666;
                color:white;
                font-family: Microsoft YaHei;
                background:#454545;
            }
            QComboBox::drop-down{
                border:none;
                border-left:1px solid #999999;
                min-width: 19px;
            }
            QComboBox::down-arrow{image: url(./photos/arrow.png);}
            QProgressBar{font-family:Microsoft YaHei;}
            QRadioButton{
                color:white;
                font-family: Microsoft YaHei;
            }
        ''')
        self.help_test.setStyleSheet('''

            QLabel{font-size:10px;font-family: Microsoft YaHei;}
        ''')

    # 鼠标可拖动无边框窗口
    def mousePressEvent(self, event):
        if (event.button() == Qt.LeftButton):
            self.m_Position = event.globalPos()-self.pos() #获取鼠标相对窗口的位置
            if (self.m_Position.y() < 60):
                self.m_flag = True
            event.accept()
    def mouseMoveEvent(self, QMouseEvent):
        if (Qt.LeftButton and self.m_flag):
            self.move(QMouseEvent.globalPos()-self.m_Position)#更改窗口位置
            QMouseEvent.accept()
    def mouseReleaseEvent(self, QMouseEvent):
        self.m_flag = False

    # 更改转换模式
    def ChangeMode(self):
        if (0 == self.pdf_combobox.currentIndex()):
            self.pdf_file_label.setText('选择PDF文件：')
            self.pdf_save_label.setText('图片保存位置：')
            self.a_flag = 1
        if (1 == self.pdf_combobox.currentIndex()):
            self.t_flag += 1
            self.pdf_file_label.setText('选择图片文件：')
            self.pdf_save_label.setText('PDF保存位置：')
            self.a_flag = 0

    # 转换主逻辑
    def Transform(self):
        # PDF转图片
        if (self.a_flag):
            if (1 == self.FileJudge('PDF') and 1 == self.PathJudge('PDF')):
                self.pdf_progressbar.setVisible(True)
                if ('' == self.pdf_file_lineedit.text()):
                    self.pdf_save_lineedit.setText(os.path.dirname(self.pdf_file_lineedit.text()))
                    self.filepath = self.pdf_save_lineedit.text().replace('\\', '/')
                    self.filename = os.path.basename(self.pdf_file_lineedit.text().replace('\\', '/'))
                for pg in range(self.doc.pageCount):
                    page = self.doc[pg]
                    rotate = int(0)
                    # 每个尺寸的缩放系数为2，这将为我们生成分辨率提高四倍的图像。
                    zoom_x = 2
                    zoom_y = 2
                    trans = fitz.Matrix(zoom_x, zoom_y).preRotate(rotate)
                    pm = page.getPixmap(matrix=trans, alpha=False)
                    pm.writePNG(self.filepath + '/' + self.filename[:-4] + '{:01}.png' .format(pg))
                    self.pdf_progressbar.setValue(int(100 * ((pg + 1) / self.doc.pageCount)))
                self.SuccessTransf()
                self.pdf_progressbar.setVisible(False)
                if (self.pdf_radiobutton.isChecked()):
                    os.startfile(self.filepath)
                self.doc.close()
        
        # 图片转PDF
        else:
            if (1 == self.FilesJudge() and 1 == self.PathJudge('PDF')):
                self.pdf_progressbar.setVisible(True)
                doc = fitz.open()
                pp = 0
                self.filepath = os.path.dirname(self.fileslist[0])
                self.filename = os.path.basename(self.fileslist[0][:-4])
                if ('' == self.pdf_file_lineedit.text()):
                    p = self.filepath
                    self.pdf_save_lineedit.setText(os.path.dirname(self.fileslist[0]))
                else:
                    p = self.pdf_save_lineedit.text()
                if os.path.exists(self.filepath + '/' + self.filename + '.pdf'):
                    os.remove(self.filepath + '/' + self.filename + '.pdf')
                for img in FilenameSort(self.fileslist): # 读取图片，确保按文件名排序
                    pp += 1
                    imgdoc = fitz.open(img)         # 打开图片
                    pdfbytes = imgdoc.convertToPDF()    # 使用图片创建单页的 PDF
                    imgpdf = fitz.open("pdf", pdfbytes)
                    doc.insertPDF(imgpdf)          # 将当前页插入文档
                    self.pdf_progressbar.setValue(int(100 * ((pp + 1) / len(self.fileslist))))
                doc.save(p + '/' + self.filename + '.pdf')          # 保存pdf文件
                doc.close()
                self.SuccessInverseTransf()
                self.pdf_progressbar.setVisible(False)
                if (self.pdf_radiobutton.isChecked()):
                    os.startfile(self.filepath)

    def W_Transform(self):
        # Word转PDF
        self.word_progressbar.setVisible(True)
        def CreatePdf(wordPath, pdfPath):
            """
            word转pdf
            :param wordPath: word文件路径
            :param pdfPath:  生成pdf文件路径
            """
            word = gencache.EnsureDispatch('Word.Application')
            doc = word.Documents.Open(wordPath, ReadOnly=1)
            doc.ExportAsFixedFormat(pdfPath,
                                    constants.wdExportFormatPDF,
                                    Item=constants.wdExportDocumentWithMarkup,
                                    CreateBookmarks=constants.wdExportCreateHeadingBookmarks)
            word.Quit(constants.wdDoNotSaveChanges)
        wordname = self.word_file_lineedit.text().replace('\\', '/')
        if ('.doc' == wordname[-4:]):
            pdfname = wordname[:-4] + '.pdf'
        if ('.docx' == wordname[-5:]):
            pdfname = wordname[:-5] + '.pdf'
        CreatePdf(wordname, pdfname)
        self.doc = fitz.open(pdfname)
        for pg in range(self.doc.pageCount):
            page = self.doc[pg]
            rotate = int(0)
            # 每个尺寸的缩放系数为4，这将为我们生成分辨率提高八倍的图像。
            zoom_x = 4.0
            zoom_y = 4.0
            trans = fitz.Matrix(zoom_x, zoom_y).preRotate(rotate)
            pm = page.getPixmap(matrix=trans, alpha=False)
            pm.writePNG(self.filepath + '/' + self.filename[:-4] + '{:01}.png' .format(pg))
            self.word_progressbar.setValue(int(100 * ((pg + 1) / self.doc.pageCount)))
        self.SuccessTransf()
        self.doc.close()
        os.remove(pdfname)
        self.word_progressbar.setVisible(False)
        if (self.word_radiobutton.isChecked()):
            os.startfile(self.word_save_lineedit.text())

    # 判断文件合法性
    def FileJudge(self, mode):
        if ('PDF' == mode):
            if ('' != self.pdf_file_lineedit.text()):
                if ('.pdf' == self.pdf_file_lineedit.text()[-4:] and os.path.exists(self.pdf_file_lineedit.text())):
                    self.doc = fitz.open(self.pdf_file_lineedit.text())
                    self.pdf_remind_label_1.setText('当前文档共 %s 页' % self.doc.pageCount)
                    return 1
                else:
                    self.WrongPathBox()
                    return 0
            else:
                return 0
        if ('Word' == mode):
            thebool = ('.doc' == self.word_file_lineedit.text()[-4:] or 'docx' == self.word_file_lineedit.text()[-4:])
            if ('' != self.word_file_lineedit.text()):
                if (thebool and os.path.exists(self.word_file_lineedit.text())):
                    #self.doc = fitz.open(self.word_file_lineedit.text())
                    #self.word_remind_label_1.setText('当前文档共 %s 页' % self.doc.pageCount)
                    return 1
                else:
                    self.WrongPathBox()
                    return 0
            else:
                return 0
    
    # 浏览文件动作
    def Openfile(self, mode):
        if ('PDF' == mode and self.a_flag):
            self.openfile_name = QtWidgets.QFileDialog.getOpenFileName(self, "浏览",'', "PDF Files (*.pdf);;All Files (*)")
            self.pdf_file_lineedit.setText(self.openfile_name[0])
            if (self.FileJudge('PDF')):
                self.filename = os.path.basename(self.pdf_file_lineedit.text().replace('\\', '/'))
                if ('' == self.pdf_save_lineedit.text()):
                    self.filepath = os.path.dirname(self.pdf_file_lineedit.text().replace('\\', '/'))
                    self.pdf_save_lineedit.setText(self.filepath)
                self.SuccessReadBox()
        if ('Word' == mode and self.a_flag):
            self.openfile_name = QtWidgets.QFileDialog.getOpenFileName(self, "浏览",'', "Word Files (*.doc; *.docx);;All Files (*)")
            self.word_file_lineedit.setText(self.openfile_name[0])
            if (self.FileJudge('Word')):
                self.filename = os.path.basename(self.word_file_lineedit.text().replace('\\', '/'))
                if ('' == self.word_save_lineedit.text()):
                    self.filepath = os.path.dirname(self.word_file_lineedit.text().replace('\\', '/'))
                    self.word_save_lineedit.setText(self.filepath)
        # 浏览多文件动作
        if (not self.a_flag):
            lineedittext = ''
            self.openfile_names = QtWidgets.QFileDialog.getOpenFileNames(self, "浏览",'', "Picture Files (*.png; *.jpg;);;All Files (*)")
            if (self.FilesJudge()):
                for i in range (len(self.fileslist)):
                    lineedittext += self.fileslist[i]
                    lineedittext += '；'
                self.pdf_file_lineedit.setText(lineedittext)
                self.pdf_remind_label_1.setText('当前共选择 %s 张图片' % len(self.fileslist))
                self.SuccessLoadPictrue()
                if ('' == self.pdf_save_lineedit.text()):
                    self.filepath = os.path.dirname(self.fileslist[0])
                    self.pdf_save_lineedit.setText(self.filepath)
    
    # 判断路径合法性
    def PathJudge(self, mode):
        if ('PDF' == mode):
            if ('' != self.pdf_save_lineedit.text):
                if (os.path.exists(self.pdf_save_lineedit.text())):
                    return 1
                else:
                    self.WrongPathRead()
                    return 0
        if ('Word' == mode):
            if ('' != self.word_save_lineedit.text):
                if (os.path.exists(self.word_save_lineedit.text())):
                    return 1
                else:
                    self.WrongPathRead()
                    return 0
    
    
    # 浏览路径动作
    def OpenPath(self, mode):
        if ('PDF' == mode):
            self.openfile_path = QtWidgets.QFileDialog.getExistingDirectory(self, "浏览")
            self.pdf_save_lineedit.setText(self.openfile_path)
            if (self.PathJudge('PDF')):
                self.filepath = self.pdf_save_lineedit.text().replace('\\', '/')
        if ('Word' == mode):
            self.openfile_path = QtWidgets.QFileDialog.getExistingDirectory(self, "浏览")
            self.word_save_lineedit.setText(self.openfile_path)
            if (self.PathJudge('Word')):
                self.filepath = self.word_save_lineedit.text().replace('\\', '/')

    # 判断多文件合法性
    def FilesJudge(self):
        self.fileslist = []                                     # 初始化多文件列表
        for i in range (len(self.openfile_names[0])):
            if ('' != self.openfile_names[0][i]):
                if (('.png' == self.openfile_names[0][i][-4:] or '.jpg' == self.openfile_names[0][i][-4:]) and os.path.exists(self.openfile_names[0][i])):
                    self.fileslist.append(self.openfile_names[0][i])
                else:
                    self.WrongFilesRead()
            else:
                continue
        if (0 == len(self.fileslist)):
            self.WrongFilesRead()
            return 0
        else:
            return 1

    # 消息弹窗
    def InterBox(self, flag, title, text):
        self.interbox_dialog = QtWidgets.QDialog()
        self.interbox_layout = QtWidgets.QVBoxLayout()
        self.interbox_dialog.setLayout(self.interbox_layout)
        self.interbox_dialog.setFixedSize(280, 120)
        self.interbox_dialog.setWindowTitle(title)

        self.interbox_label = QtWidgets.QLabel()
        if (0 == flag):
            self.interbox_label.setText(text % self.doc.pageCount)
        if (1 == flag):
            self.interbox_label.setText(text % len(self.fileslist))
        if (2 == flag):
            self.interbox_label.setText(text)

        self.interbox_okbutton = QtWidgets.QPushButton('确定')
        self.interbox_okbutton.clicked.connect(self.interbox_dialog.close)
        self.interbox_okbutton.setFixedSize(110, 26)

        self.interbox_layout.addWidget(self.interbox_label, 0, Qt.AlignCenter | Qt.AlignBottom)
        self.interbox_layout.addWidget(self.interbox_okbutton, 0, Qt.AlignBottom | Qt.AlignRight)

        # 消息弹窗样式
        self.interbox_dialog.setStyleSheet('''
            QDialog{
                background:#535353;
            }
            QLabel{
                color: white;
                font-family:Microsoft YaHei;
            }
            QPushButton{
                border-radius:13px;
                border:1px solid #bebebe;
                color: white;
                font-family: Microsoft YaHei;
                background:#535353;
            }
            QPushButton:hover{
                border-radius:13px;
                border:2px solid #3defff;
                color: white;
                font-family: Microsoft YaHei;
                background:#535353;
            }
            QPushButton:pressed{
                border:none;
                border-radius:13px;
                color: black;
                font-family: Microsoft YaHei;
                background:white;
            }
        ''')

        self.interbox_dialog.setWindowFlags(Qt.WindowCloseButtonHint | Qt.WindowStaysOnTopHint)
        self.interbox_dialog.setWindowIcon(QIcon('./photos/1.4.ico'))
        self.interbox_dialog.open()
    def SuccessTransf(self):
        self.InterBox(0, '\^o^/', '转换成功！\n共输出了 %s 张图片')
    def SuccessInverseTransf(self):
        self.InterBox(1, '\^o^/', '转换成功！\n合成了一份含有 %s 张图片的PDF')
    def SuccessReadBox(self):
        self.InterBox(0, 'q(≧▽≦q)', '读取成功！当前文档共 %s 页\n点击“转换”开始转换图片')
    def SuccessLoadPictrue(self):
        self.InterBox(1, 'q(≧▽≦q)', '成功！当前共选择 %s 张图片 \n点击“转换”开始转换成PDF')
    def WrongPathBox(self):
        self.InterBox(2, '(っ °Д °;)っ', '错误！\n请检查文件格式或路径是否正确')
    def WrongPathRead(self):
        self.InterBox(2, '(っ °Д °;)っ', '您当前选择的输出路径有误或不存在\n输出路径将自动切换为图片所在位置')
    def WrongFilesRead(self):
        self.InterBox(2, '(っ °Д °;)っ', '错误！您当前未选择任何可用的图片文件\n请重新选择')

def main():
    app = QtWidgets.QApplication(sys.argv)
    gui = MainUi()
    gui.show()
    sys.exit(app.exec_())

if __name__ == '__main__':
    main()

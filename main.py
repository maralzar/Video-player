'''
            project:video player
            programmer:Maral zarvani
            option: play video , skip to 

            Skip the video and select the section of interest using the predefined tags in Excel
            Increase or decrease the volume
            Edit tags and save them again at the selected location
            Play at any speed
            Set the theme
            Zoom in
            Define a shortcut key to go back and forth in the video and zoom in

'''
from PyQt5 import QtCore
from PyQt5 import  QtWidgets
from PyQt5.QtMultimedia import QMediaContent, QMediaPlayer
from PyQt5.QtMultimediaWidgets import QVideoWidget
from PyQt5.QtWidgets import (QApplication, QFileDialog, QHBoxLayout, QLabel,QComboBox,QTableWidget,QTableWidgetItem,
        QPushButton, QSizePolicy, QSlider, QStyle, QVBoxLayout,QMessageBox,QWidget,QMainWindow,QAction,QShortcut,QDoubleSpinBox )
import sys
import pandas as pd
from PyQt5 import QtGui
import numpy as np

class VideoWindow(QMainWindow):
    
    def __init__(self, parent=None):
        super(VideoWindow, self).__init__(parent)
        self.setWindowTitle("Video Player") 
        self.setStyleSheet("background-image : url(.//image//bg//2.png);") 
        self.value="0"
        self.path="."
        self.mediaPlayer = QMediaPlayer(None, QMediaPlayer.VideoSurface)
        videoWidget = QVideoWidget()
        videoWidget.setStyleSheet("background-color: black;")
        style="border : 1px solid black ;min-height: 20px;min-width: 20px;border-radius : 10px;"
        #play button
        self.playButton = QPushButton()
        self.playButton.setStyleSheet(style) 
        self.playButton.setEnabled(False)
        self.playButton.setIcon(self.style().standardIcon(QStyle.SP_MediaPlay))
        self.playButton.clicked.connect(self.play)
        #set speed for play 
        self.spin_speed = QDoubleSpinBox (self) 
        self.spin_speed.setStyleSheet(style)
        self.spin_speed.setDecimals(1)
        self.spin_speed.setSingleStep(0.1)
        self.spin_speed.setProperty("value", 1.0)
        self.spin_speed.valueChanged.connect(self.doublex)
        #addin tag
        self.test_btn = QPushButton("...")
        self.test_btn.setStyleSheet(style)
        self.test_btn.clicked.connect(self.addtag)
        #edit tag 
        self.edit_btn = QPushButton("Edit")
        self.edit_btn.setStyleSheet(style)
        self.edit_btn.setEnabled(False)
        self.edit_btn.clicked.connect(self.edit_tag)
        #combi box for displaying tag
        self.cb = QComboBox()
        self.cb.setStyleSheet(style+"background-image : url(.//image//icon//index.jpg);")
        self.cb.activated.connect(self.skip)

        #full screen button
        self.fullscreenButton = QPushButton()
        self.fullscreenButton.setStyleSheet(style) 
        self.fullscreenButton.setIcon(QtGui.QIcon('.//image//icon//index.png'))
        self.fullscreenButton.clicked.connect(self.show_fullscreen)
        #slider for video 
        self.positionSlider = QSlider(QtCore.Qt.Horizontal)
        self.positionSlider.setRange(0, 0)
        self.positionSlider.sliderMoved.connect(self.setPosition)
        #label for error
        self.errorLabel = QLabel("ok")
        self.errorLabel.setSizePolicy(QSizePolicy.Preferred,QSizePolicy.Maximum)
        #change theme
        default= QAction(QtGui.QIcon('.//image//icon//color.jpg'), '&Default color', self) 
        default.setStatusTip('change color')
        default.triggered.connect(self.default_color)

        theme1 = QAction(QtGui.QIcon('.//image//icon//color.jpg'), '&Sky mode', self) 
        theme1.setStatusTip('sky mode')
        theme1.triggered.connect(self.theme1)

        theme2 = QAction(QtGui.QIcon('.//image//icon//color.jpg'), '&White color', self) 
        theme2.setStatusTip('white color mode')
        theme2.triggered.connect(self.theme2)

        theme3 = QAction(QtGui.QIcon('.//image//icon//color.jpg'), '&Dark mode', self) 
        theme3.setStatusTip('dark')
        theme3.triggered.connect(self.theme3)
        #change volume
        self.sld = QSlider(QtCore.Qt.Horizontal, self)
        self.sld.setFocusPolicy(QtCore.Qt.NoFocus)
        self.sld.valueChanged.connect(self.changeValue)
        self.sld.setStyleSheet("background: red; left: 4px; right: 4px; height: 5px;"+style) 
        #volume picture
        self.label = QLabel(self)
        pixmap1 = QtGui.QPixmap(".//image//icon//volume.jpg")
        self.pixmap = pixmap1.scaled(25,25)
        self.label.setPixmap(self.pixmap)
        self.label.setFixedSize(25, 25)
        self.label.setStyleSheet("border-radius : 50;   border : 2px solid black;") 
        self.sld.valueChanged.connect(self.changeValue)
        #label for show duration
        
        # self.sld.valueChanged.connect(self.changeValue)
        # self.videoWidget.mouseDoubleClickEvent.connect(self.mouseDoubleClickEvent)
        # Create open action
        openAction = QAction(QtGui.QIcon('.//image//icon//open.png'), '&Open', self)        
        openAction.setShortcut('Ctrl+O')
        openAction.setStatusTip('Open movie')
        openAction.triggered.connect(self.openFile)

        # Create exit action
        exitAction = QAction(QtGui.QIcon('.//image//icon//exit.png'), '&Exit', self)        
        exitAction.setShortcut('Ctrl+Q')
        exitAction.setStatusTip('Exit application')
        exitAction.triggered.connect(self.exitCall)

        # Create menu bar and add action
        menuBar = self.menuBar()
        fileMenu = menuBar.addMenu('&File')
        fileMenu.addAction(openAction)
        fileMenu.addAction(exitAction)
        #Theme change
        ViewMenu = menuBar.addMenu('&View')
        ViewMenu.addAction(default)
        ViewMenu.addAction(theme1)
        ViewMenu.addAction(theme2)
        ViewMenu.addAction(theme3)
        

        # Create a widget for window contents
        wid = QWidget(self)
        self.setCentralWidget(wid)

        # Create layouts to place inside widget
        controlLayout = QHBoxLayout()

        #place items for  first layout
        controlLayout.setContentsMargins(0, 0, 0, 0) #set cordinate
        controlLayout.addWidget(self.playButton)
        controlLayout.addWidget(self.spin_speed)
        controlLayout.addWidget(self.cb)
        controlLayout.addWidget(self.test_btn)
        controlLayout.addWidget(self.edit_btn)
        controlLayout.addWidget(self.sld)
        controlLayout.addWidget(self.label)
        controlLayout.addWidget(self.fullscreenButton)
        controlLayout.addWidget(self.errorLabel)
        
       
       #place item for second layout
        layout = QVBoxLayout()
        layout.addWidget(videoWidget)
        layout.addWidget(self.positionSlider)
        
        layout.addLayout(controlLayout)
       # Set widget to contain window contents
        wid.setLayout(layout)

        self.mediaPlayer.setVideoOutput(videoWidget)
        self.mediaPlayer.stateChanged.connect(self.mediaStateChanged)
        self.mediaPlayer.positionChanged.connect(self.positionChanged)
        self.mediaPlayer.durationChanged.connect(self.durationChanged)
        self.mediaPlayer.error.connect(self.handleError)

        #keyboard shortcut
        #key right forward 1000 
        self.shortcut = QShortcut(QtGui.QKeySequence(QtCore.Qt.Key_Right), self)
        self.shortcut.activated.connect(self.forwardSlider)
        self.shortcut = QShortcut(QtGui.QKeySequence(QtCore.Qt.Key_Left), self)
        self.shortcut.activated.connect(self.backSlider)
        self.shortcut = QShortcut(QtGui.QKeySequence(QtCore.Qt.Key_Escape), self)
        self.shortcut.activated.connect(self.exit_fullscreen)
    #set value for slider 
    def changeValue(self, value):
        self.mediaPlayer.setVolume(value)
        
    def onValueChanged(self, val):
        print(val)
     #play video with selected speed
    def doublex(self):
        value = self.spin_speed.value() 
        self.mediaPlayer.setPlaybackRate(value)
    #add tag  for video
    def addtag(self):

        fileName, _ = QFileDialog.getOpenFileName(self, "Open Tag",QtCore.QDir.homePath())
        
        if fileName != '':
            print(fileName)
            df = pd.read_excel(fileName) #read excel file from your selected directory
            s = df.values 
            print(s)               #store list of list of value for editing    
            self.value = s
            self.path = fileName           #store path for save edited excel file
            self.edit_btn.setEnabled(True)  #active btn because tag was load and we can edit it
            for i in s:
                self.cb.addItem(i[1])       #load combo box
     
    def edit_tag(self):
        ex = EditMessageBox (self.value) #Create Object and sent value for fill the table for editing
        a = ex.r() 
        
        if not isinstance(a,np.ndarray):         #call fuction to return modified data
            self.value = a.values  
            self.cb.clear()
            for i in a.values:  #update combobox
                    self.cb.addItem(i[1]) 
            
            #save edited value to current address
            a.to_excel(self.path,index=0,header = None, engine='xlsxwriter')
            
    #forward the video for 1000*60 ms
    def forwardSlider(self):
        self.mediaPlayer.setPosition(self.mediaPlayer.position() + 1000*60)  

    #bachward the video for 1000*60 ms
    def backSlider(self):
        self.mediaPlayer.setPosition(self.mediaPlayer.position() - 1000*60)

# DOUBLE click the full screen enabled
    def mouseDoubleClickEvent(self, event):
        if self.isFullScreen():
            self.showNormal()
        else:
            self.showFullScreen()
       
    def default_color(self):
        self.setStyleSheet("background-image : url(.//image//bg//2.png);")  

    def theme1(self):
        self.setStyleSheet("background-image : url(.//image//bg//5.jpg);") 
    
    def theme2(self):
        self.setStyleSheet("background-image : url(.//image//bg//4.jpg);") 
    def theme3(self):
        self.setStyleSheet("background-color: gray;") 
        #
    def openFile(self):
        fileName, _ = QFileDialog.getOpenFileName(self, "Open Movie",
                QtCore.QDir.homePath())

        if fileName != '':
            self.mediaPlayer.setMedia(
                    QMediaContent(QtCore.QUrl.fromLocalFile(fileName)))
            self.playButton.setEnabled(True)
            
            
        # exit from application
    def exitCall(self):
        sys.exit(app.exec_())

    def play(self):
        if self.mediaPlayer.state() == QMediaPlayer.PlayingState:
       
            self.mediaPlayer.pause()
           

        else:
            self.mediaPlayer.play()
            self.mediaPlayer.setVolume(self.sld.value())
            self.errorLabel.setText("ok")

    def mediaStateChanged(self, state):
        print("mediaStateChanged")
        if self.mediaPlayer.state() == QMediaPlayer.PlayingState:
            self.playButton.setIcon(
                    self.style().standardIcon(QStyle.SP_MediaPause))
        else:
            self.playButton.setIcon(
                    self.style().standardIcon(QStyle.SP_MediaPlay))

    def positionChanged(self, position):
      
        self.positionSlider.setValue(position)

    def durationChanged(self, duration):
        print("duration",duration)
        self.positionSlider.setRange(0, duration)
    #skiping to our determined time
    def skip(self):
        print(self.value)
        index = self.cb.currentIndex()
        t = self.value[index][0]
        self.setPosition(int(t)*1000)
        self.positionSlider.setValue(int(t)*1000)

    

    def setPosition(self, position):
        print("setPosition")
        self.mediaPlayer.setPosition(position)
        # handle error when video cant load or some thing has cased error
    def handleError(self):
        self.playButton.setEnabled(False)
        self.errorLabel.setText("Error: " + self.mediaPlayer.errorString())
        #show in full screen
    def show_fullscreen(self):
        msg = QMessageBox()
        QMessageBox.about(self, " ", "برای خروج دکمه escرا فشار دهید")
        
        # retval = msg.exec_()
        self.showFullScreen()
        self.fullscreenButton.setIcon(QtGui.QIcon('.//image//icon//outzoom.png'))

    def exit_fullscreen(self):
        self.showNormal()


class EditMessageBox(QtWidgets.QMessageBox):
    def __init__(self,items):
        QtWidgets.QMessageBox.__init__(self)
        self.setSizeGripEnabled (True)

        self.setWindowTitle('ویرایش تگ ها')
        self.setIcon(self.Question)
        self.setText("ویرایش")
        self.a=items
        self.addButton (
            QtWidgets.QPushButton('ذخیره'), 
            QtWidgets.QMessageBox.YesRole
        )
    
        self.addButton(
            QtWidgets.QPushButton('لغو'), 
            QtWidgets.QMessageBox.RejectRole
        )

        self.addTableWidget (self,items)

        currentClick = self.exec_()

        if currentClick==0 :
            rows = self.tableWidget.rowCount()
            columns = self.tableWidget.columnCount() 
            df = pd.DataFrame()
            for i in range(rows):
                for j in range(columns):
                    df.loc[i, j] = str(self.tableWidget.item(i, j).text())  
            self.a=df    

          
           
        if currentClick==2 :
            pass
            

    def addTableWidget (self, parentItem,items) :
        self.l =  QtWidgets.QVBoxLayout()
        self.tableWidget = QtWidgets.QTableWidget(parentItem)
        self.tableWidget.setObjectName ('tableWidget')
       

        self.tableWidget.setColumnCount(2)
        self.tableWidget.setRowCount(len(items))
        self.tableWidget.move(30,80)
        self.tableWidget.resize(500, 170)
        self.tableWidget.setHorizontalHeaderLabels(['ثانیه', 'موضوع'])
        for i  in range(0,len(items)):
            self.tableWidget.setItem(i,0,QTableWidgetItem(str(items[i][0])))
            self.tableWidget.setItem(i,1,QTableWidgetItem(str(items[i][1]))) 

        self.l.addWidget(self.tableWidget)
        self.setLayout(self.l)

    def event(self, e):
        result = QtWidgets.QMessageBox.event(self, e)
        self.setMinimumWidth(0)
        self.setMaximumWidth(16777215)
        self.setMinimumHeight(0)
        self.setMaximumHeight(16777215)
        self.setSizePolicy(
            QtWidgets.QSizePolicy.Expanding, 
            QtWidgets.QSizePolicy.Expanding
        )
        self.resize(550, 300)
        return result 
    def r(self):
        return self.a
               
        
if __name__ == '__main__':
    app = QApplication(sys.argv)
    app.setStyle("Fusion")
    player = VideoWindow()
    player.resize(640, 480)
    player.show()
    sys.exit(app.exec_())
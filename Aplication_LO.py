# -*- coding: utf-8 -*-

# Form implementation generated from reading ui file 'Interface.ui'
#
# Created by: PyQt5 UI code generator 5.11.3
#
# WARNING! All changes made in this file will be lost!
import sys
import os
import shutil


#Katlog instalacyjny QGIS
QGIS_PATH = r'C:\Program Files\QGIS 3.4\apps\qgis-ltr'
#Ustawienie sciezek
sys.path.append( os.path.join(QGIS_PATH, 'python') )
os.environ['PATH'] = '{};{};{}'.format(os.path.join(QGIS_PATH, 'bin'), os.path.join(QGIS_PATH, '../qt5/bin'), os.environ['PATH'])
os.environ['QT_PLUGIN_PATH'] = '{};{}'.format(os.path.join(QGIS_PATH, 'qtplugins'), os.path.join(QGIS_PATH, '../qt5/plugins'))
os.environ['QGIS_PREFIX_PATH'] = QGIS_PATH
from PyQt5 import QtCore, QtGui, QtWidgets
from PyQt5.QtCore import QSize, QTimer,Qt,QEventLoop, QVariant,QSettings,QFileInfo
from PyQt5.QtWidgets import QDialog,QLabel,QHBoxLayout,QSizePolicy, QAction, QFileDialog, QMessageBox, QColorDialog, QTableView
from PyQt5.QtGui import QImage, QColor, QPainter,QFont, QCursor, QPixmap, QIcon,QTextBlock
from PyQt5.QtPrintSupport import QPrintDialog, QPrinter
from qgis.core import *
from qgis.gui import *

import inspect
import types
import xlwt
from xlrd import open_workbook
from os import path
import xlsxwriter
from zipfile import ZipFile 
import zipfile


  
#from PyQt5.QtGui import QTableView

from qgis.gui import QgsAttributeTableModel, QgsAttributeTableView, QgsAttributeTableFilterModel



QgsApplication.setPrefixPath(QGIS_PATH, True)

  


class Ui_MainWindow(object):
   
   
   
    def setupUi(self, MainWindow):
    
        selectedLayer=None
        qgisInstance=None
        categoryDictionary=None
        listOfFieldsNames=None
        iloscWszWydz=None
        calPowLes=None
        lista2=[]
        lista4=[]
        dataZIP=None
       
        MainWindow.setObjectName("MainWindow")
        MainWindow.resize(1144, 593)
        MainWindow.setWindowTitle("LASY OCHRONNE")
        icon = QtGui.QIcon()
        icon.addPixmap(QtGui.QPixmap(":/images/leaf.ico"), QtGui.QIcon.Normal, QtGui.QIcon.Off)
        MainWindow.setWindowIcon(icon)
        self.centralwidget = QtWidgets.QWidget(MainWindow)
        self.centralwidget.setObjectName("centralwidget")
        
        
        statusbar = QtWidgets.QStatusBar(MainWindow)
        statusbar.setGeometry(QtCore.QRect(0, 935, 1900, 60))
        statusbar.setObjectName("statusbar")
        
        self.label2 = QtWidgets.QLabel(statusbar)
       
        self.label2.setGeometry(QtCore.QRect(50, 10, 250, 25))
        self.label2.setObjectName("label12")
        self.label2.setText("<html><head/><body><p><span style=\" font-size:9pt; font-weight:600;\">WSPOLRZEDNE:</span></p></body></html>")
        self.label2.setVisible(False)
        
        
        self.label144 = QtWidgets.QLineEdit(statusbar)
        self.label144.setGeometry(QtCore.QRect(200, 10, 450, 25))
        self.label144.setReadOnly(True)
       
       
        self.label144.setObjectName("label144")
        self.label144.setVisible(False)
  
        
        self.iloscWybranychWydzielen = QtWidgets.QLabel(statusbar)
        self.iloscWybranychWydzielen.setGeometry(QtCore.QRect(680, 10, 620, 25))
        self.iloscWybranychWydzielen.setObjectName("iloscWybranychWydzielen")
       
        
        self.powierzchniaWybranychWydzielen = QtWidgets.QLabel(statusbar)
        self.powierzchniaWybranychWydzielen.setGeometry(QtCore.QRect(1350, 10, 550, 25))
        self.powierzchniaWybranychWydzielen.setObjectName("powierzchniaWybranychWydzielen")
       
        
        

        self.progressBar = QtWidgets.QProgressBar(statusbar)
        self.progressBar.setGeometry(QtCore.QRect(1780, 10, 118, 23))
        self.progressBar.setProperty("value", 24)
        self.progressBar.setObjectName("progressBar")
        self.progressBar.setVisible(False)
        
        self.groupBox12 = QtWidgets.QGroupBox(self.centralwidget)
        self.groupBox12.setGeometry(QtCore.QRect(20, 20, 550, 120))
        self.groupBox12.setObjectName("groupBox12")
        self.groupBox12.setTitle("Wczytaj dane")
        self.loadLayerBtn = QtWidgets.QPushButton(self.groupBox12)
        self.loadLayerBtn.setGeometry(QtCore.QRect(40, 40, 201, 41))
        self.loadLayerBtn.setStyleSheet("border-image: url(:/images/wczytaj.png);")
        self.loadLayerBtn.setText("")
        self.loadLayerBtn.setObjectName("loadLayerBtn")
        self.deleteLayerBtn= QtWidgets.QPushButton(self.groupBox12)
        self.deleteLayerBtn.setGeometry(QtCore.QRect(260, 40, 201, 41))
        self.deleteLayerBtn.setStyleSheet("border-image: url(:/images/resetuj.png);")
        self.deleteLayerBtn.setText("")
        self.deleteLayerBtn.setObjectName("deleteLayerBtn")
        self.groupBox_2 = QtWidgets.QGroupBox(self.centralwidget)
        self.groupBox_2.setGeometry(QtCore.QRect(20, 155, 550, 111))
        self.groupBox_2.setObjectName("groupBox_2")
        self.groupBox_2.setTitle("Zaznacz lasy ochronne")
        self.comboBox5 = QtWidgets.QComboBox(self.groupBox_2)
        self.comboBox5.setGeometry(QtCore.QRect(30, 40, 220, 41))
        self.comboBox5.setObjectName("comboBox5")
        font = QtGui.QFont()
        font.setPointSize(7)
        self.comboBox5.setFont(font)
        
        self.zaznaczBtn = QtWidgets.QPushButton(self.groupBox_2)
        self.zaznaczBtn.setGeometry(QtCore.QRect(270, 40, 121, 41))
        self.zaznaczBtn.setStyleSheet("border-image: url(:/images/zaznacz.png);")
        self.zaznaczBtn.setText("")
        self.zaznaczBtn.setObjectName("zaznaczBtn")
        self.odznaczBtn = QtWidgets.QPushButton(self.groupBox_2)
        self.odznaczBtn.setGeometry(QtCore.QRect(400, 40, 121, 41))
        self.odznaczBtn.setStyleSheet("border-image: url(:/images/odznacz.png);")
        self.odznaczBtn.setText("")
        self.odznaczBtn.setObjectName("odznaczBtn")
        
        self.groupBox_3 = QtWidgets.QGroupBox(self.centralwidget)
        self.groupBox_3.setGeometry(QtCore.QRect(20, 295, 550, 380))
        self.groupBox_3.setObjectName("groupBox_3")
        self.groupBox_3.setTitle("Nadaj styl")
        
        
        self.mColorButton = QgsColorButton()
        self.mColorButton.setObjectName("mColorButton")
        self.mColorButton.setColor(QColor(255,255,255))
        
        self.mColorButton_2 = QgsColorButton()
        self.mColorButton_2.setObjectName("mColorButton_2")
        self.mColorButton_2.setColor(QColor(74,149,51))
       
        self.mColorButton_3 = QgsColorButton()
        self.mColorButton_3.setObjectName("mColorButton_3")
        self.mColorButton_3.setColor(QColor(0,0,0))
        
        self.mColorButton_4 = QgsColorButton()
        self.mColorButton_4.setObjectName("mColorButton_4")
        self.mColorButton_4.setColor(QColor(255,255,255))
       
        self.mColorButton_5 = QgsColorButton()
        self.mColorButton_5.setObjectName("mColorButton_5")
        self.mColorButton_5.setColor(QColor(162,162,162))
      
        self.treeWidget = QtWidgets.QTreeWidget(self.groupBox_3)
        self.treeWidget.setGeometry(QtCore.QRect(30, 60, 490, 265))
        self.treeWidget.setObjectName("treeWidget")
        
       
       
    
        
        self.treeWidget.setAlternatingRowColors(True)
        self.treeWidget.setHeaderHidden(True)
        
        self.treeWidget.header().resizeSection(0, 250)
   
        self.treeWidget.setObjectName("treeWidget")
        item_0 = QtWidgets.QTreeWidgetItem(self.treeWidget)
        item_0.setExpanded(True)
        item_1 = QtWidgets.QTreeWidgetItem(item_0)
        self.treeWidget.setItemWidget(item_1,1,self.mColorButton)
        item_0 = QtWidgets.QTreeWidgetItem(self.treeWidget)
        item_0.setExpanded(True)
        item_1 = QtWidgets.QTreeWidgetItem(item_0)
        self.treeWidget.setItemWidget(item_1,1,self.mColorButton_2)
        item_1 = QtWidgets.QTreeWidgetItem(item_0)
        self.treeWidget.setItemWidget(item_1,1,self.mColorButton_3)
        item_0 = QtWidgets.QTreeWidgetItem(self.treeWidget)
        item_0.setExpanded(True)
        item_1 = QtWidgets.QTreeWidgetItem(item_0)
        self.treeWidget.setItemWidget(item_1,1,self.mColorButton_4)
        item_1 = QtWidgets.QTreeWidgetItem(item_0)
        self.treeWidget.setItemWidget(item_1,1,self.mColorButton_5)
        
      
    
    
        self.treeWidget.setAlternatingRowColors(True)
        self.treeWidget.setHeaderHidden(True)
       
        __sortingEnabled = self.treeWidget.isSortingEnabled()
        self.treeWidget.setSortingEnabled(False)
        self.treeWidget.headerItem().setText(0,  "1")
        self.treeWidget.headerItem().setText(1, "2")
       
        self.treeWidget.topLevelItem(0).setText(0, "MAPA")
        self.treeWidget.topLevelItem(0).child(0).setText(0,  "Kolor tla wydruku")
        self.treeWidget.topLevelItem(1).setText(0,  "LASY OCHRONNE")
        self.treeWidget.topLevelItem(1).child(0).setText(0, "Kolor wypelnienia")
        self.treeWidget.topLevelItem(1).child(1).setText(0,  "Kolor konturu")
        self.treeWidget.topLevelItem(2).setText(0,"LASY POZOSTALE")
        self.treeWidget.topLevelItem(2).child(0).setText(0,  "Kolor wypelnienia")
        self.treeWidget.topLevelItem(2).child(1).setText(0,  "Kolor konturu")
       
       
        
        self.groupBox_4 = QtWidgets.QGroupBox(self.centralwidget)
        self.groupBox_4.setGeometry(QtCore.QRect(20, 700, 550, 141))
        self.groupBox_4.setObjectName("groupBox_4")
        self.groupBox_4.setTitle("Generuj raport")
        self.pushButton_5 = QtWidgets.QPushButton(self.groupBox_4)
        self.pushButton_5.setGeometry(QtCore.QRect(70, 40, 421, 71))
        self.pushButton_5.setStyleSheet("border-image: url(:/images/raport.png);")
        self.pushButton_5.setText("")
        self.pushButton_5.setObjectName("pushButton_5")
        self.mColorButton.setAllowOpacity(True)
        self.mColorButton.setBehavior(QgsColorButton.SignalOnly)
        self.mColorButton.setShowNoColor(True)
        self.mColorButton_2.setAllowOpacity(True)
        self.mColorButton_2.setBehavior(QgsColorButton.SignalOnly)
        self.mColorButton_2.setShowNoColor(True)
        self.mColorButton_3.setAllowOpacity(True)
        self.mColorButton_3.setBehavior(QgsColorButton.SignalOnly)
        self.mColorButton_3.setShowNoColor(True)
        self.mColorButton_4.setAllowOpacity(True)
        self.mColorButton_4.setBehavior(QgsColorButton.SignalOnly)
        self.mColorButton_4.setShowNoColor(True)
        self.mColorButton_5.setAllowOpacity(True)
        self.mColorButton_5.setBehavior(QgsColorButton.SignalOnly)
        self.mColorButton_5.setShowNoColor(True)
        
        
        
        
        
        
        self.map = QgsMapCanvas(self.centralwidget)
        self.map.setGeometry(QtCore.QRect(600, 30, 1290, 810))
        self.map.setObjectName("map")
      
        
        MainWindow.setCentralWidget(self.centralwidget)
        self.menubar = QtWidgets.QMenuBar(MainWindow)
        self.menubar.setGeometry(QtCore.QRect(0, 0, 1144, 18))
        self.menubar.setObjectName("menubar")
        self.menuProjekt = QtWidgets.QMenu(self.menubar)
        self.menuProjekt.setTitle("Projekt")
        self.menuProjekt.setObjectName("menuProjekt")
        self.menuOpcje = QtWidgets.QMenu(self.menubar)
        self.menuOpcje.setObjectName("menuOpcje")
        self.menuOpcje.setTitle("Opcje")
        self.menuPomoc = QtWidgets.QMenu(self.menubar)
        self.menuPomoc.setObjectName("menuPomoc")
        
        MainWindow.setMenuBar(self.menubar)
        self.toolBar = QtWidgets.QToolBar(MainWindow)
        self.toolBar.setObjectName("toolBar")
        MainWindow.addToolBar(QtCore.Qt.TopToolBarArea, self.toolBar)
        self.statusBar = QtWidgets.QStatusBar(MainWindow)
        self.statusBar.setObjectName("statusBar")
        MainWindow.setStatusBar(self.statusBar)
        self.actionNowy_projekt = QtWidgets.QAction(MainWindow)
        icon1 = QtGui.QIcon()
        icon1.addPixmap(QtGui.QPixmap(":/images/document.png"), QtGui.QIcon.Normal, QtGui.QIcon.Off)
        self.actionNowy_projekt.setIcon(icon1)
        self.actionNowy_projekt.setObjectName("actionNowy_projekt")
        self.actionZapisz_jako_obraz = QtWidgets.QAction(MainWindow)
        icon2 = QtGui.QIcon()
        icon2.addPixmap(QtGui.QPixmap(":/images/save.png"), QtGui.QIcon.Normal, QtGui.QIcon.Off)
        self.actionZapisz_jako_obraz.setIcon(icon2)
        self.actionZapisz_jako_obraz.setObjectName("actionZapisz_jako_obraz")
        self.actionDrukuj_obraz = QtWidgets.QAction(MainWindow)
        icon3 = QtGui.QIcon()
        icon3.addPixmap(QtGui.QPixmap(":/images/print.png"), QtGui.QIcon.Normal, QtGui.QIcon.Off)
        self.actionDrukuj_obraz.setIcon(icon3)
        self.actionDrukuj_obraz.setObjectName("actionDrukuj_obraz")
        
       
        
        actionIdentify= QtWidgets.QAction(MainWindow)
        icon6 = QtGui.QIcon()
        icon6.addPixmap(QtGui.QPixmap(":/images/identify.svg"), QtGui.QIcon.Normal, QtGui.QIcon.Off)
        actionIdentify.setIcon(icon6)
        actionIdentify.setObjectName("Identify")
        
        actionPokazTabeleAtrybutow = QtWidgets.QAction(MainWindow)
        icon7 = QtGui.QIcon()
        icon7.addPixmap(QtGui.QPixmap(":/images/openTabele.svg"), QtGui.QIcon.Normal, QtGui.QIcon.Off)
        actionPokazTabeleAtrybutow.setIcon(icon7)
      
        actionPokazTabeleAtrybutow.setObjectName("actionPokazTabeleAtrybutow")
        
        
    
        

        self.actionEksportujTabeleAtrybutow = QtWidgets.QAction(MainWindow)
        icon8 = QtGui.QIcon()
        icon8.addPixmap(QtGui.QPixmap(":/images/exportToExcel.png"), QtGui.QIcon.Normal, QtGui.QIcon.Off)
        self.actionEksportujTabeleAtrybutow.setIcon(icon8)
        self.actionEksportujTabeleAtrybutow.setObjectName("actionEksportujTabeleAtrybutow")
        
        
        
        self.actionHelp = QtWidgets.QAction(MainWindow)
        icon4 = QtGui.QIcon()
        icon4.addPixmap(QtGui.QPixmap(":/images/help.png"), QtGui.QIcon.Normal, QtGui.QIcon.Off)
        self.actionHelp.setIcon(icon4)
        self.actionHelp.setObjectName("actionHelp")
        
        
        self.actionPan = QtWidgets.QAction(MainWindow)
        icon5 = QtGui.QIcon()
        icon5.addPixmap(QtGui.QPixmap(":/images/pan"), QtGui.QIcon.Normal, QtGui.QIcon.Off)
        
        self.actionPan.setIcon(icon5)
        self.actionPan.setObjectName("PAN")
        self.actionPan.setCheckable(True)
    
        
        
        self.actionZoomIn = QtWidgets.QAction(MainWindow)
        icon9 = QtGui.QIcon()
        icon9.addPixmap(QtGui.QPixmap(":/images/zoom_in.png"), QtGui.QIcon.Normal, QtGui.QIcon.Off)
        self.actionZoomIn.setIcon(icon9)
        self.actionZoomIn.setObjectName("ZoomIn")
        
        
        self.actionZoomOut= QtWidgets.QAction(MainWindow)
        icon10 = QtGui.QIcon()
        icon10.addPixmap(QtGui.QPixmap(":/images/zoom_out.png"), QtGui.QIcon.Normal, QtGui.QIcon.Off)
        self.actionZoomOut.setIcon(icon10)
        self.actionZoomOut.setObjectName("ZoomOut")
        

        self.actionZoomIn.setCheckable(True)
        self.actionZoomOut.setCheckable(True)
        actionIdentify.setCheckable(True)
        
        
        self.actionCaly_zasieg = QtWidgets.QAction(MainWindow)
        icon12 = QtGui.QIcon()
        icon12.addPixmap(QtGui.QPixmap(":/images/zoom_full.png"), QtGui.QIcon.Normal, QtGui.QIcon.Off)
        self.actionCaly_zasieg.setIcon(icon12)
        self.actionCaly_zasieg.setObjectName("actionCaly_zasieg")
      
        self.menuProjekt.addAction(self.actionNowy_projekt)
        self.menuProjekt.addAction(self.actionZapisz_jako_obraz)
        self.menuProjekt.addAction(self.actionDrukuj_obraz)
        self.menuOpcje.addAction(actionIdentify)
        self.menuOpcje.addAction(actionPokazTabeleAtrybutow)
        self.menuOpcje.addAction(self.actionEksportujTabeleAtrybutow)
    
       
        self.menuPomoc.addAction(self.actionHelp)
        self.menuPomoc.setTitle("Pomoc")
        self.toolBar.setWindowTitle("toolBar")
        self.menubar.addAction(self.menuProjekt.menuAction())
        self.menubar.addAction(self.menuOpcje.menuAction())
        self.menubar.addAction(self.menuPomoc.menuAction())
        
        self.toolBar.addAction(self.actionNowy_projekt)
        self.toolBar.addAction(self.actionZapisz_jako_obraz)
        self.toolBar.addAction(self.actionDrukuj_obraz)
        self.toolBar.addSeparator()
        self.toolBar.addAction(self.actionPan)
      
        self.toolBar.addAction(self.actionZoomIn)
        self.toolBar.addAction(self.actionZoomOut)
 
        self.toolBar.addAction(self.actionCaly_zasieg)
        
        self.toolBar.addSeparator()
        
        
        self.toolBar.addAction(actionIdentify)
        self.toolBar.addAction(actionPokazTabeleAtrybutow)
        self.toolBar.addAction(self.actionEksportujTabeleAtrybutow)
        
        self.toolBar.addSeparator()
        self.toolBar.addAction(self.actionHelp)
        
        self.actionNowy_projekt.setText("Nowy projekt")
        self.actionZapisz_jako_obraz.setText("Zapisz jako obraz")
        self.actionDrukuj_obraz.setText("Drukuj obraz")
        
        actionPokazTabeleAtrybutow.setText("Pokaz tabele wlasciwosci wydzielen")
        self.actionPan.setText("Przesun")
        self.actionEksportujTabeleAtrybutow.setText("Eksportuj tabele wlasciwosci wydzielen do pliku Excel")
        actionIdentify.setText("Pokaz wlasciwosci wydzielenia")
        self.actionZoomIn.setText("Powieksz")
        self.actionZoomOut.setText("Pomniejsz")
    
        self.actionCaly_zasieg.setText("Caly zasieg")
        self.actionHelp.setText("Podrecznik Uzytkownika")
       
        
        colorFillSelectedWydz=self.mColorButton_2.color().name()
        colorBorderSelectedWydz=self.mColorButton_3.color().name()
        colorFillOtherWydz=self.mColorButton_4.color().name()
        colorBorderOthertWydz=self.mColorButton_5.color().name()
        bacgroundColor=self.mColorButton.color()
        self.map.setCanvasColor(bacgroundColor)
        
        
        self.toolPan = QgsMapToolPan(self.map)
        self.toolPan.setAction(self.actionPan)
        
        self.toolZoomIn = QgsMapToolZoom(self.map, False) # false = in
        self.toolZoomIn.setAction(self.actionZoomIn)
        self.toolZoomOut = QgsMapToolZoom(self.map, True) # true = out
        self.toolZoomOut.setAction(self.actionZoomOut)
        
        

        
        def createCategory():
            nonlocal categoryDictionary
            
            
            symbolsCategory=['OCH','OCH GLEB', 'OCH WOD', 'OCH USZK','OCH BADAW','OCH NAS','OCH OSTOJ','OCH MIAST','OCH UZDR','OCH OBR']
            categories=["wszystkie kategorie ochronnosci", 
                    "glebochronne",
                    "wodochronne", 
                    "trwale uszkodzone przemyslowo",
                    "stale powierzchnie doswiadczalne",
                    "nasienne",
                    "ostoje zwierzat", 
                    "w miastach i wokol miast", 
                    "uzdrowiskowe",
                    "obronne"]
           
            
            for category in categories:
                categoryDictionary={categories[n]: symbolsCategory[n] for n in range(len(categories))}
    
        
            
            self.comboBox5.addItems(categories)
            self.comboBox5.setCurrentIndex(-1)
        
        
        def addStyleLayer(map):
            global selectedLayer
            selectedLayerSymbol = QgsFillSymbol.createSimple({'color':'#ffffff','color_border':'#a2a2a2'})
            renderer = QgsSingleSymbolRenderer(selectedLayerSymbol)
            selectedLayer.setRenderer(renderer) 
            map.refresh()
            
        def changeFieldNames():
            global selectedLayer
            global listOfFieldsNames
            oldNames = selectedLayer.fields().names()
            newNames=["NUMER POWIERZCHNI", "ADRES LESNY","RODZAJ POWIERZCHNI", "TYP SIEDLISKOWY LASU", "GOSPODARSTWO", "FUNKCJA LASU", "BUDOWA PIONOWA DRZEWOSTANU", "WIEK REBNOSCI", "POWIERZCHNIA (ha)", "KATEGORIE OCHRONNOSCI", "KOD GATUNKU PANUJACEGO", "UDZIAL GATUNKU PANUJACEGO", "WIEK GATUNKU PANUJACEGO", "ROK STANU DANYCH"]
            dictionaryOfoldAndnewFields = dict(zip(oldNames , newNames))
            selectedLayer.startEditing()
            for oldName, newName in dictionaryOfoldAndnewFields.items():
                for field in  selectedLayer.fields():
                    if field.name() == oldName:
                        idx =  selectedLayer.fields().indexFromName(field.name())
                        selectedLayer.renameAttribute(idx, newName)
            selectedLayer.updateFields()
            listOfFieldsNames=selectedLayer.fields().names()
            
            
        def loadLayer( map ):
            """Ladowanie warstwy"""
            global selectedLayer
            global qgisInstance
            global symbol1
            global symbol2
            global iloscWszWydz
            global calPowLes
            global dataZIP
            
            #Wybranie pliku
            #selectedShapefile = QFileDialog.getOpenFileName(None, "Wybierz plik shp", "", "Shapefile (*.shp)")
            dialog = QFileDialog()
            dialog.setFileMode(QFileDialog.DirectoryOnly)
            selectedShapefile1 = dialog.getOpenFileName(None, "Wybierz archiwum", "", "Archiwum ZIP (*.zip)")
            y=str(os.path.basename(selectedShapefile1[0])).replace('.zip','/')
            print(y)
            x=selectedShapefile1[0]
            
            # opening the zip file in READ mode 
            with ZipFile(x, 'r') as zip:
                desktop = os.path.expanduser("~/Desktop")
                zip.extractall(desktop)
                dataZIP =desktop+'/'+y+'G_SUBAREA.shp'
               
            #dataZIP=str(desktop+'/'+y.replace('/',''))
            print(dataZIP)
            selectedLayer = QgsVectorLayer(dataZIP, "Wydzielenia", "ogr")
            crs = QgsCoordinateReferenceSystem('EPSG:2180')
            selectedLayer.setCrs(crs)
            #Dodane warstw do mapy
            qgisInstance=QgsProject.instance()
            qgisInstance.addMapLayers([selectedLayer])
            map.setExtent(selectedLayer.extent()) 
            addStyleLayer(map)
            map.setLayers([selectedLayer])
            changeFieldNames()
            createCategory()
            iloscWszWydz=selectedLayer.featureCount()
            calPowLes=round(sum(feature["POWIERZCHNIA (ha)"]  for feature in selectedLayer.getFeatures()),2)
            self.label144.setVisible(True)
            self.label2.setVisible(True)
            self.map.xyCoordinates.connect(lambda:showXY(self.label144, self.map.mouseLastXY()))
            
        
        
        def clearLabel():
            self.iloscWybranychWydzielen.setText("")
            self.powierzchniaWybranychWydzielen.setText("")
            
  
        def deleteLayer(map):
            """ Usuwanie warstwy"""
            global selectedLayer
            global qgisInstance
            global dataZIP
            qgisInstance.removeAllMapLayers()
            map.refresh()
            clearLabel()
            self.comboBox5.setCurrentIndex(-1)
            self.map.xyCoordinates.disconnect()
            self.label144.setVisible(False)
            self.label2.setVisible(False)
          
            
        def createSymbol(colorFill,colorBorder):
            
            symbol=QgsFillSymbol.createSimple({'color': colorFill, 'color_border':colorBorder})
            return symbol
        
        
        def setSymbol(map):
            global selectedLayer
            renderer = QgsCategorizedSymbolRenderer()
            renderer.setClassAttribute('KATEGORIE OCHRONNOSCI')      
            for value in lista2:
               
                symbol=createSymbol(colorFillSelectedWydz,colorBorderSelectedWydz)
                renderer.addCategory(QgsRendererCategory(value, symbol, 'ab'))
         
                
            for value in lista4:
                symbol=createSymbol(colorFillOtherWydz,colorBorderOthertWydz)
               
                renderer.addCategory(QgsRendererCategory(value, symbol, 'cb'))   
           
            selectedLayer.setRenderer(renderer) 
            map.refresh()
        
            
        def odznaczWydzielenia(map):
            nonlocal lista2
            nonlocal lista4
            addStyleLayer(map)
            clearLabel()
            self.comboBox5.setCurrentIndex(-1)
            lista2=[]
            lista4=[]
           
        def wybierzWydzielenia(map):
            #global selectedFeature
            global selectedLayer
            
            nonlocal colorFillSelectedWydz
            nonlocal colorBorderSelectedWydz
            nonlocal colorFillOtherWydz
            nonlocal colorBorderOthertWydz
            nonlocal categoryDictionary
            global iloscWszWydz
            global calPowLes
            nonlocal lista2
            nonlocal lista4
            
            
                   
            lista=[]
            lista3=[]
            lista5=[]
            categories=[]
          
             
            
            for category in categoryDictionary.keys():
                if category==self.comboBox5.currentText():
                    for feat in selectedLayer.getFeatures():
                        if str(categoryDictionary[category]) in str(feat["KATEGORIE OCHRONNOSCI"]):
                            lista.append(feat["KATEGORIE OCHRONNOSCI"])
                            lista2=set(lista)
                            lista5.append(feat["POWIERZCHNIA (ha)"])
                        else:
                            lista3.append(feat["KATEGORIE OCHRONNOSCI"])
                            lista4=set(lista3)
                    
                           
        
            setSymbol(map)
           
            
            iloscWybWydz=len(lista)
            powZazLasOchr=round(sum(lista5),2)
           
        
            self.iloscWybranychWydzielen.setText(f"<html><head/><body><p><span style=\" font-size:9pt; font-weight:600;\">Wybrano {iloscWybWydz} wydzielen lasow ochronnych z {iloscWszWydz} wszystkich wydzielen</span></p></body></html>")
            self.powierzchniaWybranychWydzielen.setText(f"<html><head/><body><p><span style=\" font-size:9pt; font-weight:600;\">Powierzchnia zaznaczonych lasow ochronnych: {powZazLasOchr} ha</span></p></body></html>")
            
            
        class GISDialog(QDialog):
            def __init__(self, title,icon, isContextHelpButtonHint,isMSWindowsFixedSizeDialogHint, parent=None):
                self.title=title
                self.icon=icon
                self.isContextHelpButtonHint=isContextHelpButtonHint
                self.isMSWindowsFixedSizeDialogHint=isMSWindowsFixedSizeDialogHint
                super(GISDialog, self).__init__(parent)
                self.setWindowFlag(Qt.WindowContextHelpButtonHint, isContextHelpButtonHint)
                self.setWindowFlag(Qt.MSWindowsFixedSizeDialogHint, isMSWindowsFixedSizeDialogHint)
                self.setWindowFlags(Qt.WindowStaysOnTopHint)
                self.setWindowTitle(title)
                self.setWindowIcon(QIcon(icon))
        
        
                                
            def closeEvent(self, event):
               
                global selectedLayer      
                selectedLayer.removeSelection()
               
              
        class openDialog(GISDialog):
       
            def closeEvent(self, event):
                global selectedLayer
                nonlocal actionPokazTabeleAtrybutow
                nonlocal actionIdentify
                actionPokazTabeleAtrybutow.setEnabled(True) 
                actionIdentify.setEnabled(True)    
                selectedLayer.removeSelection()
           
            def showEvent(self, event):
                nonlocal actionPokazTabeleAtrybutow
                nonlocal actionIdentify
                actionPokazTabeleAtrybutow.setEnabled(False)
                actionIdentify.setEnabled(False)  
                
                
                 
            
        class CustomSelectTool(QgsMapToolIdentify):   

            def __init__(self, canvas):
                QgsMapToolIdentify.__init__(self, canvas)
                self.canvas = canvas
                self.featureId = None
                

        
            def canvasReleaseEvent(self, event):
                
                global selectedLayer
                global listOfFieldsNames
                dlg = GISDialog("Wlasciwosci wydzielenia",':/images/info.png', False, True )
                
                found_features = self.identify(event.x(), event.y(), self.TopDownStopAtFirst, self.VectorLayer)
                if len(found_features) > 0:
                    layer = found_features[0].mLayer
                    feature = found_features[0].mFeature
                    selectedLayer.selectByIds([feature.id()]) 
                    geometry = feature.geometry()
                   
                    
                   
                    layout=QHBoxLayout()
                    tableWidget = QtWidgets.QTableWidget()
                    
                    #tableWidget.setGeometry(QtCore.QRect(170, 20, 256, 192))
                    
                    tableWidget.setObjectName("tableWidget")
                   
                        
                    tableWidget.setColumnCount(2)
                    tableWidget.setRowCount(len(listOfFieldsNames))
        
                    font = QtGui.QFont()
                    font.setBold(True)
                    font.setWeight(75)
                    
                    item = QtWidgets.QTableWidgetItem()
                    item.setFlags( Qt.ItemIsSelectable |  Qt.ItemIsEnabled )
                    item.setFont(font)
                    tableWidget.setHorizontalHeaderItem(0, item)
                    item = QtWidgets.QTableWidgetItem()
                    item.setFlags( Qt.ItemIsSelectable |  Qt.ItemIsEnabled )
                    item.setFont(font)
                    tableWidget.setHorizontalHeaderItem(1, item)
                    
                    
                   

                    item = tableWidget.horizontalHeaderItem(0)
                    item.setText("Wlasciwosc")
                    item = tableWidget.horizontalHeaderItem(1)
                    item.setText("Wartosc")
                    
                    
                
                                 
              
                    for field in listOfFieldsNames:
                        attribute = feature.attribute(field)
                       
                        item = QtWidgets.QTableWidgetItem()
                        item.setFlags(Qt.ItemIsEnabled )
                        
                        tableWidget.setItem(listOfFieldsNames.index(field), 0, item)
                        
                        item = QtWidgets.QTableWidgetItem()
                        item.setFlags(Qt.ItemIsEnabled )
                        tableWidget.setItem(listOfFieldsNames.index(field), 1, item)
                        
                        item = tableWidget.item(listOfFieldsNames.index(field), 0)
                        item.setText(str(field))
                    
                        item = tableWidget.item(listOfFieldsNames.index(field), 1)
                        item.setText(str(attribute))
                        
                      
                    
                        
                       
                    
                    tableWidget.setSizeAdjustPolicy(QtWidgets.QAbstractScrollArea.AdjustToContents)
                    tableWidget.resizeColumnsToContents()
                   
                    layout.addWidget(tableWidget)
                    
                    
                    dlg.setLayout(layout)
                   
                    dlg.exec_()
                 
                    
        toolInfo=CustomSelectTool(self.map)
        toolInfo.setAction(actionIdentify)     
        
         
        def identify():
         
                      
            self.map.setMapTool(toolInfo)
            setCanvasCursor(':/images/cursor_arrow_INFO.svg')
        
    
                
       
            


       
        def showXY(label,p):
            # SLOT. Show coordinates
            global selectedLayer
            tool=QgsMapTool(self.map)
            #self.map.setMapTool(tool)
            point= tool.toLayerCoordinates(selectedLayer, p)
            coords = "  X: "+ str(point.x()) +"/ Y: "+str( point.y())
            
            label.setText(coords)
           
            fm=label.fontMetrics()
            w = fm.boundingRect(coords).width();
            label.setFixedWidth(w);
              
        
                            
            
        def openTabele(map):
            
            global selectedLayer
            selectedLayer.getFeatures()
            listOfFeature=[]
           
           
            pan()
            tableView = QTableView()
          
                
            vector_layer_cache = QgsVectorLayerCache(selectedLayer, 10000)
            attribute_table_model = QgsAttributeTableModel(vector_layer_cache)
           
            attribute_table_model.loadLayer()
            attribute_table_filter_model = QgsAttributeTableFilterModel(map, attribute_table_model)
           
            tableView.setModel(attribute_table_filter_model)
            tableView.resizeColumnsToContents()
            selectionModel = tableView.selectionModel()
            
            def changeSelectionEvent():
                nonlocal listOfFeature
                listOfFeature.clear()
                listOfFeature=[attribute_table_filter_model.rowToId(rowID) for rowID in tableView.selectedIndexes()]
                selectedLayer.selectByIds(listOfFeature)
               
            
            selectionModel.selectionChanged.connect(lambda:changeSelectionEvent())
            
            dlg =openDialog("Tabela wlasciwosci wydzielen",':/images/openTabele.svg', False,False)
           
            dlg.resize(800,800)     
            layout=QHBoxLayout()
            layout.addWidget(tableView)
            dlg.setLayout(layout)
            dlg.setModal(False)
            dlg.show()
          
           
            loop=QEventLoop()
            loop.exec_()
            
          
        def extent(map):
            global  selectedLayer   
            map.setExtent(selectedLayer.extent())
            map.refresh()
            
        def help():
            os.startfile("./help/Instrukcja_uzytkownika.pdf")
            
        def zoomIn():
            self.map.setMapTool(self.toolZoomIn)
            setCanvasCursor(':/images/cursor_zoomIn.svg')
            
        def zoomOut():
            self.map.setMapTool(self.toolZoomOut)
            setCanvasCursor(':/images/cursor_zoomOut.svg')


        def setCanvasCursor(btm):
            pm = QPixmap(btm)
            bm = pm.createMaskFromColor(QColor(0xFF, 0x00, 0xF0), Qt.MaskOutColor)
            cursor=QCursor(pm)
            self.map.setCursor(cursor)
            
        def pan():
            self.map.setMapTool(self.toolPan)
            
  
                
        def setColor(button):
            color = QColorDialog.getColor(options=QColorDialog.ShowAlphaChannel)
        
            button.setColor(color)
            
            
        def setBackgroudMapColor(map):
            
            
            bacgroundColor=self.mColorButton.color()
            map.setCanvasColor(bacgroundColor)
            map.refresh()
            
        def printImage():
            global selectedLayer 
            printer = QPrinter(QPrinter.HighResolution)
            dialog = QPrintDialog(printer)
            if dialog.exec_() == QPrintDialog.Accepted:
                
            
                
                options = QgsMapSettings()
                options.setLayers([selectedLayer])
                options.setBackgroundColor(QColor(255, 255, 255))
                options.setOutputSize(QSize(1500, 800))
                options.setExtent(selectedLayer.extent())
                render = QgsMapRendererParallelJob(options)
                render.print_(printer)
                
        def saveImage( map ):
            """ Zapisywanie obrazu"""
            global selectedLayer
            global qgisInstance
            image_types = "JPeg (*.jpg);;Bitmap(*.bmp);;PNG (*.png);;Tiff (*.tiff)"
            options = QFileDialog.Options()
            fileName = QFileDialog.getSaveFileName(None, "Zapisz jako obraz", "", filter=image_types,options=options)
            #selectedLayer = QgsVectorLayer(selectedShapefile[0], "Wybrana warstwa", "ogr")
            if fileName:
               
               
                
                    
                project = QgsProject.instance()
            
                        
                layout = QgsPrintLayout(project)
                layout.initializeDefaults()
               
                
                 
                # create map item in the layout
                map = QgsLayoutItemMap(layout)
                map.setRect(20, 20, 20, 20)
                 
                # set the map extent
                ms = QgsMapSettings()
                ms.setLayers([selectedLayer]) # set layers to be mapped
                rect = QgsRectangle(ms.fullExtent())
                rect.scale(1.0)
                ms.setExtent(rect)
                map.setExtent(rect)
                map.setBackgroundColor(QColor(255, 255, 255, 0))
                layout.addLayoutItem(map)
                 
                map.attemptMove(QgsLayoutPoint(5, 20, QgsUnitTypes.LayoutMillimeters))
                map.attemptResize(QgsLayoutSize(180, 180, QgsUnitTypes.LayoutMillimeters))
                 
                legend = QgsLayoutItemLegend(layout)
                legend.setTitle("Legend")
                layerTree = QgsLayerTree()
                layerTree.addLayer(selectedLayer)
                legend.model().setRootGroup(layerTree)
                layout.addLayoutItem(legend)
                legend.attemptMove(QgsLayoutPoint(230, 15, QgsUnitTypes.LayoutMillimeters))
                 
                #scalebar = QgsLayoutItemScaleBar(layout)
                #scalebar.setStyle('Line Ticks Up')
                #scalebar.setUnits(QgsUnitTypes.DistanceKilometers)
                #scalebar.setNumberOfSegments(4)
                #scalebar.setNumberOfSegmentsLeft(0)
                
                #scalebar.setLinkedMap(map)
                #scalebar.setUnitLabel('km')
                #scalebar.setFont(QFont('Arial', 14))
                #scalebar.update()
                #layout.addLayoutItem(scalebar)
                #scalebar.attemptMove(QgsLayoutPoint(220, 190, QgsUnitTypes.LayoutMillimeters))
                
                scaleBar = QgsLayoutItemScaleBar(layout)
                #scaleBar.setUnits(QgsUnitTypes.DistanceKilometers)
                scaleBar.setLinkedMap(map)
                scaleBar.applyDefaultSettings()
                scaleBar.applyDefaultSize()
                
                #scaleBar.setUnitLabel('km')
                #scaleBar.setUnitsPerSegment(0.05)
               
                # scaleBar.setStyle('Line Ticks Down') 
                scaleBar.setNumberOfSegmentsLeft(0)
                scaleBar.setNumberOfSegments (4)
                scaleBar.update()
            
                
                # item.update()
                layout.addItem(scaleBar)
                scaleBar.attemptMove(QgsLayoutPoint(220, 190, QgsUnitTypes.LayoutMillimeters))
                
               
                
                title = QgsLayoutItemLabel(layout)
                title.setText("Lasy ochronne")
                title.setFont(QFont('Arial', 24))
                title.adjustSizeToText()
                layout.addLayoutItem(title)
                title.attemptMove(QgsLayoutPoint(10, 5, QgsUnitTypes.LayoutMillimeters))
                title.setFrameEnabled(True) 
               

                exporter = QgsLayoutExporter(layout)
                 
                
                exporter = QgsLayoutExporter(layout)
                exporter.exportToImage(fileName[0], QgsLayoutExporter.ImageExportSettings())
                                
                                  
     
                                
        
        self.loadLayerBtn.clicked.connect(lambda:loadLayer(self.map))
        self.deleteLayerBtn.clicked.connect(lambda:deleteLayer(self.map))
        
        self.actionCaly_zasieg.triggered.connect(lambda:extent(self.map))
        #self.actionHelp.triggered.connect(lambda:help())
       
      
        
        self.zaznaczBtn.clicked.connect(lambda:wybierzWydzielenia(self.map))
        
       
        
        self.actionDrukuj_obraz.triggered.connect(lambda:printImage())
        #self.actionZapisz_jako_obraz.triggered.connect(lambda:saveImage( self.map ))
        self.pushButton_5.clicked.connect(lambda:(generateRaport()))
        self.odznaczBtn.clicked.connect(lambda:odznaczWydzielenia(self.map))
        actionPokazTabeleAtrybutow.triggered.connect(lambda:openTabele(self.map))
        self.actionEksportujTabeleAtrybutow.triggered.connect(lambda:excelSave())
        self.mColorButton.colorChanged.connect(lambda:setBackgroudMapColor(self.map))
        self.mColorButton_2.colorChanged.connect(lambda:colorsChange(self.map))
        self.mColorButton_3.colorChanged.connect(lambda:colorsChange(self.map))
        self.mColorButton_4.colorChanged.connect(lambda:colorsChange(self.map))
        self.mColorButton_5.colorChanged.connect(lambda:colorsChange(self.map))
        self.mColorButton.clicked.connect(lambda:setColor(self.mColorButton))
        self.mColorButton_2.clicked.connect(lambda:setColor(self.mColorButton_2))
        self.mColorButton_3.clicked.connect(lambda:setColor(self.mColorButton_3))
        self.mColorButton_4.clicked.connect(lambda:setColor(self.mColorButton_4))
        self.mColorButton_5.clicked.connect(lambda:setColor(self.mColorButton_5))
        self.actionPan.triggered.connect(lambda:pan())
        self.actionZoomIn.triggered.connect(lambda:zoomIn())
        self.actionZoomOut.triggered.connect(lambda:zoomOut())
        actionIdentify.triggered.connect(lambda:identify())
        
  
  
        
        QtCore.QMetaObject.connectSlotsByName(MainWindow)
        
        
        def colorsChange(map):
            nonlocal colorFillSelectedWydz
            nonlocal colorBorderSelectedWydz
            nonlocal colorFillOtherWydz
            nonlocal colorBorderOthertWydz
            nonlocal lista2
            nonlocal lista4
            global selectedLayer
            
          
            colorFillSelectedWydz=self.mColorButton_2.color().name()
            colorBorderSelectedWydz=self.mColorButton_3.color().name()
            colorFillOtherWydz=self.mColorButton_4.color().name()
            colorBorderOthertWydz=self.mColorButton_5.color().name()
             
            setSymbol(map)
            
        def excelSave():
            global selectedLayer
            
           
            names = selectedLayer.fields().names()
    
            dirPath = QSettings().value("/excelSavePath", ".", type=str)    
            (filename, filter) = QFileDialog.getSaveFileName(MainWindow,
                        "Zapisz jako plik excel...",
                        dirPath,
                        "Excel files (*.xls)",
                        "Filter list for selecting files from a dialog box")
            fn, fileExtension = path.splitext(filename)
            if len(fn) == 0: # user choose cancel
                return
            QSettings().setValue("/excelSavePath", QFileInfo(filename).absolutePath())
            if fileExtension != '.xls':
                filename = filename + '.xls'
    
            
            
            class Writer:
    
                fileName = '/tmp/example2.xls'
    
                wb = None
                ws = None
    
                def __init__(self, filename):
                    self.fileName = filename
                    self.wb = xlwt.Workbook()
                    self.ws = self.wb.add_sheet('Qgis Attributes')
            
                def writeAttributeRow(self, rowNr, attributes):
                    colNr = 0
                    for cell in attributes:
                        # QGIS2.0 does not have QVariants anymore, only for <2.0:
                        cell = str(cell)
                        try:
                            cell = float(cell)
                        except:
                            pass
            
                        self.ws.write(rowNr, colNr, cell)
                        colNr = colNr + 1
    
                def saveFile(self):
                    self.wb.save(self.fileName)
    
            
            xlw = Writer(filename)
           
            
            feature = QgsFeature();
    
            xlw.writeAttributeRow(0, names)
    
            rowNr = 1
           
            prov = selectedLayer.getFeatures()
            while prov.nextFeature(feature):
                    # attribute values, either for all or only for selection
                
                values = []
                for field in names:
                    values.append(feature.attribute(field))
                xlw.writeAttributeRow(rowNr, values)
                rowNr += 1
            xlw.saveFile()
            QMessageBox.information(MainWindow, "Sukces", "Zapisywanie do pliku .xls zakonczone sukcesem")
        
        def generateRaport():
            
            global selectedLayer
            nonlocal categoryDictionary
            Dane=[]
            
            dirPath = QSettings().value("/excelSavePath", ".", type=str)    
            (filename, filter) = QFileDialog.getSaveFileName(MainWindow,
                        "Zapisz jako plik excel...",
                        dirPath,
                        "Excel files (*.xls)",
                        "Filter list for selecting files from a dialog box")
            fn, fileExtension = path.splitext(filename)
            if len(fn) == 0: # user choose cancel
                return
            QSettings().setValue("/excelSavePath", QFileInfo(filename).absolutePath())
            if fileExtension != '.xls':
                filename = filename + '.xls'

            workbook = xlsxwriter.Workbook(filename)
            worksheet = workbook.add_worksheet()
            worksheet.write('A3', 'KATEGORIA OCHRONNOSCI')
            worksheet.write('B3', 'LICZBA WYDZIELEN')
            worksheet.write('C3', 'POWIERZCHNIA WYDZIELEN (HA)')
            worksheet.write('D3', 'LICZBA WYDZIELEN')
            worksheet.write('E3', 'POWIERZCHNIA WYDZIELEN (HA)')
            

            for category in categoryDictionary.keys():               
                lista1=[feat["POWIERZCHNIA (ha)"] for feat in selectedLayer.getFeatures()if str(categoryDictionary[category]) in str(feat["KATEGORIE OCHRONNOSCI"]) and feat["WIEK REBNOSCI"]<=40]
                lista2=[feat["POWIERZCHNIA (ha)"] for feat in selectedLayer.getFeatures()if str(categoryDictionary[category]) in str(feat["KATEGORIE OCHRONNOSCI"]) and feat["WIEK REBNOSCI"]>40]
                dane=[category,len(lista1),round(sum(lista2)),len(lista2),round(sum(lista2))]
                Dane.append(dane)
            
            row = 3
            col = 0
                    
            for category, a, b, c, d in Dane:
                worksheet.write(row, col, category)
                worksheet.write(row, col + 1, a)
                worksheet.write(row, col + 2, b)
                worksheet.write(row, col + 3, c)
                worksheet.write(row, col + 4, d)
                row += 1
    
            workbook.close() 
         
         
import resources_rc

if __name__ == "__main__":
    import sys
    class Window(QtWidgets.QMainWindow):
        def __init__(self):

            super().__init__()
        
        def closeEvent(self, event):
            
            QgsApplication.closeAllWindows()
            
            
    app = QgsApplication([], True)
    app.initQgis()
  
    MainWindow = Window()
    ui = Ui_MainWindow()
    ui.setupUi(MainWindow)
    MainWindow.showMaximized()

 
    
    sys.exit( app.exec_() )
    app.exitQgis()


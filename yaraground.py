# -*- coding: utf-8 -*-

# TODO
## Clearing Table ======> Done
## Checkbox Code ======> Done
## Stylesheet ======> Done
## Delete Tempfiles (tmprule.yar, result.txt) ======> Done
## Configuration ======> Done
## Whole Table cells Copy and Paste ======> Done
## Export table to Excel (csv) ======> Done
## Add tag option code
## Log File - working
## Key Event Handling : Hotkey
## PyInstaller

import os
import io
import csv
try:
    from PyQt5 import QtCore, QtGui, QtWidgets, Qt
except:
    print("Could not import PyQt5 for file identification. Use: pip3 install pyqt5")
try:
    import configparser
except:
    print("Could not import configparser for file identification. Use: pip3 install configparser")

try:
    import qdarkgraystyle
except:
    print("Could not import qdarkgraystyle for file identification. Use: pip3 install qdarkgraystyle")

conf_filename = 'yaraground.conf'
config = configparser.ConfigParser()

class Ui_yaraground(object):
    def setupUi(self, yaraground):
        yaraground.setObjectName("yaraground")
        yaraground.resize(1061, 633)
        yaraground.setFocusPolicy(QtCore.Qt.StrongFocus)
        #yaraground.setWindowFlags(QtCore.Qt.FramelessWindowHint)
        # Setup configuration
        self.yaraground_config = {'lastRulePath':''}
        try: 
            if len(config.read(conf_filename)) <= 0:
                #print ("nothing in it")
                config.add_section('YARAGROUND')
                config.set('YARAGROUND','lastRulePath','')
                config.set('YARAGROUND','lastTargetDIR','')
                with open(conf_filename,'w') as configfile:
                    config.write(configfile)
            
            # Initialize Configuration dictionary variable
            self.configuration = {'lastRulePath':'', 'lastTargetDIR':''}
            self.configuration['lastRulePath'] = config['YARAGROUND']['lastRulePath']
            self.configuration['lastTargetDIR'] = config['YARAGROUND']['lastTargetDIR']
        except:
            pass

        #Grid layout setting
        self.gridLayout = QtWidgets.QGridLayout(yaraground)
        self.gridLayout.setContentsMargins(11, 5, 11, 11)
        self.gridLayout.setSpacing(10)
        self.gridLayout.setObjectName("gridLayout")

        #TargetDirectory Layout
        self.TargetDIRLayout = QtWidgets.QVBoxLayout()
        self.TargetDIRLayout.setSizeConstraint(QtWidgets.QLayout.SetNoConstraint)
        self.TargetDIRLayout.setSpacing(15)
        self.TargetDIRLayout.setObjectName("TargetDIRLayout")
        self.targetDIRbutton = QtWidgets.QPushButton(yaraground)
        sizePolicy = QtWidgets.QSizePolicy(QtWidgets.QSizePolicy.Minimum, QtWidgets.QSizePolicy.Preferred)
        sizePolicy.setHorizontalStretch(0)
        sizePolicy.setVerticalStretch(0)
        sizePolicy.setHeightForWidth(self.targetDIRbutton.sizePolicy().hasHeightForWidth())
        self.targetDIRbutton.setSizePolicy(sizePolicy)
        font = QtGui.QFont()
        font.setBold(True)
        font.setWeight(75)
        self.targetDIRbutton.setFont(font)
        self.targetDIRbutton.setObjectName("targetDIRbutton")
        self.TargetDIRLayout.addWidget(self.targetDIRbutton)
        self.dirPathEditor = QtWidgets.QLineEdit(yaraground)
        self.dirPathEditor.setObjectName("dirPathEditor")
        self.dirPathEditor.setText(self.configuration['lastTargetDIR'])
        self.TargetDIRLayout.addWidget(self.dirPathEditor)

        #QTableWidget
        self.tableWidget = QtWidgets.QTableWidget(yaraground)
        self.tableWidget.setDragEnabled(True)
        self.tableWidget.setSelectionMode(QtWidgets.QAbstractItemView.ExtendedSelection)
        self.tableWidget.setColumnCount(2)
        self.tableWidget.setObjectName("tableWidget")
        self.tableWidget.resizeColumnsToContents()
        self.tableWidget.resizeRowsToContents()
        self.tableWidget.setFocusPolicy(QtCore.Qt.StrongFocus)
        self.tableWidget.installEventFilter(self.tableWidget)
        item = QtWidgets.QTableWidgetItem()
        item.setTextAlignment(QtCore.Qt.AlignCenter)
        font = QtGui.QFont()
        font.setBold(True)
        font.setWeight(75)
        item.setFont(font)
        self.tableWidget.setHorizontalHeaderItem(0, item)
        item = QtWidgets.QTableWidgetItem()
        item.setTextAlignment(QtCore.Qt.AlignCenter)
        font = QtGui.QFont()
        font.setBold(True)
        font.setWeight(75)
        item.setFont(font)
        self.tableWidget.setHorizontalHeaderItem(1, item)
        self.tableWidget.horizontalHeader().setDefaultSectionSize(200)
        self.tableWidget.horizontalHeader().setStretchLastSection(True)
        self.tableWidget.verticalHeader().setDefaultSectionSize(20)
        self.tableWidget.verticalHeader().setMinimumSectionSize(0)
        self.TargetDIRLayout.addWidget(self.tableWidget)
        self.gridLayout.addLayout(self.TargetDIRLayout, 0, 2, 1, 1)

        self.rightSideLayout = QtWidgets.QHBoxLayout()
        self.rightSideLayout.setSizeConstraint(QtWidgets.QLayout.SetFixedSize)
        self.rightSideLayout.setSpacing(6)
        self.rightSideLayout.setObjectName("rightSideLayout")

        self.scanButton = QtWidgets.QPushButton(yaraground)
        sizePolicy = QtWidgets.QSizePolicy(QtWidgets.QSizePolicy.Expanding, QtWidgets.QSizePolicy.Fixed)
        sizePolicy.setHorizontalStretch(0)
        sizePolicy.setVerticalStretch(0)
        sizePolicy.setHeightForWidth(self.scanButton.sizePolicy().hasHeightForWidth())
        self.scanButton.setSizePolicy(sizePolicy)
        font = QtGui.QFont()
        font.setPointSize(15)
        font.setBold(True)
        font.setWeight(75)
        self.scanButton.setFont(font)
        self.scanButton.setObjectName("scanButton")
        self.rightSideLayout.addWidget(self.scanButton)
        self.recursiveCheckBox = QtWidgets.QCheckBox(yaraground)
        sizePolicy = QtWidgets.QSizePolicy(QtWidgets.QSizePolicy.Fixed, QtWidgets.QSizePolicy.Fixed)
        sizePolicy.setHorizontalStretch(0)
        sizePolicy.setVerticalStretch(0)
        sizePolicy.setHeightForWidth(self.recursiveCheckBox.sizePolicy().hasHeightForWidth())
        self.recursiveCheckBox.setSizePolicy(sizePolicy)
        self.recursiveCheckBox.setMinimumSize(QtCore.QSize(50, 0))
        self.recursiveCheckBox.setObjectName("recursiveCheckBox")
        self.recursiveCheckBox.setChecked(True)
        self.rightSideLayout.addWidget(self.recursiveCheckBox)
        self.printMatchRadioButton = QtWidgets.QCheckBox(yaraground)
        sizePolicy = QtWidgets.QSizePolicy(QtWidgets.QSizePolicy.Fixed, QtWidgets.QSizePolicy.Fixed)
        sizePolicy.setHorizontalStretch(0)
        sizePolicy.setVerticalStretch(0)
        sizePolicy.setHeightForWidth(self.printMatchRadioButton.sizePolicy().hasHeightForWidth())
        self.printMatchRadioButton.setSizePolicy(sizePolicy)
        self.printMatchRadioButton.setObjectName("printMatchRadioButton")
        self.rightSideLayout.addWidget(self.printMatchRadioButton)
        
        #clear button
        self.clearButton = QtWidgets.QPushButton(yaraground)
        self.clearButton.setObjectName("clearButton")
        self.rightSideLayout.addWidget(self.clearButton)
        
        #excel button
        self.excelButton = QtWidgets.QPushButton(yaraground)
        self.excelButton.setObjectName("excelButton")
        self.rightSideLayout.addWidget(self.excelButton)

        #clipboard button
        self.clipboardButton = QtWidgets.QPushButton(yaraground)
        self.clipboardButton.setObjectName("clipboardButton")
        self.rightSideLayout.addWidget(self.clipboardButton)

        #gridlayout addlayout
        self.gridLayout.addLayout(self.rightSideLayout, 1, 2, 1, 1)
        
        #Left layout (yararule)
        self.frame = QtWidgets.QFrame(yaraground)
        self.frame.setEnabled(True)
        sizePolicy = QtWidgets.QSizePolicy(QtWidgets.QSizePolicy.Minimum, QtWidgets.QSizePolicy.Expanding)
        sizePolicy.setHorizontalStretch(0)
        sizePolicy.setVerticalStretch(0)
        sizePolicy.setHeightForWidth(self.frame.sizePolicy().hasHeightForWidth())
        self.frame.setSizePolicy(sizePolicy)
        self.frame.setMinimumSize(QtCore.QSize(500, 620))
        self.frame.setStyleSheet("")
        self.frame.setFrameShape(QtWidgets.QFrame.NoFrame)
        self.frame.setFrameShadow(QtWidgets.QFrame.Plain)
        self.frame.setLineWidth(0)
        self.frame.setObjectName("frame")
        self.RuleEditorLayout = QtWidgets.QHBoxLayout(self.frame)
        self.RuleEditorLayout.setContentsMargins(10, 10, 10, 10) # (left, upper, right, bottom)
        self.RuleEditorLayout.setSpacing(6)
        self.RuleEditorLayout.setObjectName("RuleEditorLayout")
        self.verticalLayout_2 = QtWidgets.QVBoxLayout()
        self.verticalLayout_2.setSizeConstraint(QtWidgets.QLayout.SetNoConstraint)
        self.verticalLayout_2.setSpacing(20)
        self.verticalLayout_2.setObjectName("verticalLayout_2")

        self.loadRuleButton = QtWidgets.QPushButton(self.frame)
        sizePolicy = QtWidgets.QSizePolicy(QtWidgets.QSizePolicy.Minimum, QtWidgets.QSizePolicy.Fixed)
        sizePolicy.setHeightForWidth(self.loadRuleButton.sizePolicy().hasHeightForWidth())
        self.loadRuleButton.setSizePolicy(sizePolicy)
        self.loadRuleButton.setMinimumSize(QtCore.QSize(300, 0))
        font = QtGui.QFont()
        font.setPointSize(13)
        font.setBold(True)
        font.setWeight(75)
        font.setStrikeOut(False)
        font.setKerning(True)
        font.setStyleStrategy(QtGui.QFont.PreferAntialias)
        self.loadRuleButton.setFont(font)
        self.loadRuleButton.setAutoDefault(False)
        self.loadRuleButton.setDefault(True)
        self.loadRuleButton.setFlat(False)
        self.loadRuleButton.setObjectName("loadRuleButton")
        self.verticalLayout_2.addWidget(self.loadRuleButton)

        # Yararule path
        self.yararulePath = QtWidgets.QLineEdit(self.frame)
        self.yararulePath.setObjectName("yararulePath")
        self.verticalLayout_2.addWidget(self.yararulePath)
        if len(self.configuration['lastRulePath']) > 0:
            self.yararulePath.setText(self.configuration['lastRulePath'])
            self.fp = open(self.configuration['lastRulePath'], 'r')
            self.content = self.fp.read()
        else:
            pass

        self.ruleEditorBox = QtWidgets.QPlainTextEdit(self.frame)
        self.ruleEditorBox.setMaximumSize(QtCore.QSize(16777215, 16777215))
        self.ruleEditorBox.setOverwriteMode(False)
        self.ruleEditorBox.setObjectName("ruleEditorBox")
        try:
            self.ruleEditorBox.setPlainText(self.content)
        except:
            pass
        self.verticalLayout_2.addWidget(self.ruleEditorBox)
        
        # Save Rule Button
        self.saveRuleButton = QtWidgets.QPushButton(self.frame)
        self.saveRuleButton.setObjectName("saveRuleButton")
        self.verticalLayout_2.addWidget(self.saveRuleButton)

        # Save As Button
        self.saveAsButton = QtWidgets.QPushButton(self.frame)
        self.saveAsButton.setObjectName("saveAsButton")
        self.verticalLayout_2.addWidget(self.saveAsButton)

        self.RuleEditorLayout.addLayout(self.verticalLayout_2)
        self.gridLayout.addWidget(self.frame, 0, 0, 2, 2)

        self.retranslateUi(yaraground)
        QtCore.QMetaObject.connectSlotsByName(yaraground)
        yaraground.setTabOrder(self.loadRuleButton, self.ruleEditorBox)
        yaraground.setTabOrder(self.ruleEditorBox, self.targetDIRbutton)
        yaraground.setTabOrder(self.targetDIRbutton, self.dirPathEditor)
        yaraground.setTabOrder(self.dirPathEditor, self.scanButton)
        yaraground.setTabOrder(self.scanButton, self.recursiveCheckBox)
        yaraground.setTabOrder(self.recursiveCheckBox, self.printMatchRadioButton)

        self.targetDIRbutton.clicked.connect(self.targetDIR)
        self.scanButton.clicked.connect(self.scan)
        self.clearButton.clicked.connect(self.clearTable)
        self.excelButton.clicked.connect(self.exportExcel)
        self.loadRuleButton.clicked.connect(self.loadRule)
        self.saveRuleButton.clicked.connect(self.saveRule)
        self.saveAsButton.clicked.connect(self.saveAs)
        self.clipboardButton.clicked.connect(self.clipboard)
        
    def retranslateUi(self, yaraground):
        _translate = QtCore.QCoreApplication.translate
        yaraground.setWindowTitle(_translate("yaraground", "YARAGROUND"))
        self.targetDIRbutton.setText(_translate("yaraground", "Target Directory"))
        self.tableWidget.setSortingEnabled(True)
        item = self.tableWidget.horizontalHeaderItem(0)
        item.setText(_translate("yaraground", "Rulename"))
        item = self.tableWidget.horizontalHeaderItem(1)
        item.setText(_translate("yaraground", "Filename"))
        self.scanButton.setText(_translate("yaraground", "SCAN"))
        self.recursiveCheckBox.setText(_translate("yaraground", "Recursive"))
        self.printMatchRadioButton.setText(_translate("yaraground", "Print Offset (X)"))
        self.loadRuleButton.setText(_translate("yaraground", "Load Yararule"))
        self.saveRuleButton.setText(_translate("yaraground", "Save"))
        self.saveAsButton.setText(_translate("yaraground", "Save As"))
        self.clearButton.setText(_translate("yaraground", "Clear"))
        self.excelButton.setText(_translate("yaraground", "Excel"))
        self.clipboardButton.setText(_translate("yaraground", "Clipboard"))
        self.clipboardButton.setShortcut(_translate("yaraground", "Meta+C, Ctrl+C"))

    def saveAs(self):
        content = self.ruleEditorBox.toPlainText()
        filename = QtWidgets.QFileDialog.getSaveFileName(None, "Save As", self.configuration['lastRulePath'])[0]
        print (filename)
        self.configuration['lastRulePath'] = filename
        config.set('YARAGROUND','lastRulePath', filename)

        #Write new rule
        try:
            new_content = open(filename, 'w')
            new_content.write(content)
            self.yararulePath.setText(filename)
        except:
            pass

    def saveRule(self):
        msgBox = QtWidgets.QMessageBox(None)
        answer = msgBox.question(None, '', 'Would you like to save?', msgBox.Yes | msgBox.No)

        if answer == msgBox.No:
            return
        else:
            print ("save") 
            content = self.ruleEditorBox.toPlainText()
            filename = self.yararulePath.text()

            #backup
            original_content = open(filename, 'r').read()
            with open(filename+".backup", 'w') as original_rule:
                original_rule.write(original_content)

            #Write new rule
            try:
                new_content = open(filename, 'w')
                new_content.write(content)
                self.configuration['lastRulePath'] = filename
                config.set('YARAGROUND','lastRulePath', filename)
                msgBox.setText("Saved")
                msgBox.exec_()
            except:
                msgBox.setText("Fail to Save")
                msgBox.exec_()


    def clearTable(self):
        self.tableWidget.setRowCount(0)
        return

    def clipboard(self):
        result = ""
        for selectionRange in self.tableWidget.selectedRanges():
            print (selectionRange)
        try:
            startrow = selectionRange.topRow()
            endrow = selectionRange.bottomRow()+1

            for row in range(startrow, endrow):
                rulename_cell = self.tableWidget.item(row, 0).text()
                filename_cell = self.tableWidget.item(row, 1).text()
                result += rulename_cell + " " + filename_cell

            QtWidgets.QApplication.clipboard().setText(result)
        except:
            return

    def exportExcel(self):
        try:
            filename = QtWidgets.QFileDialog.getSaveFileName()[0]

            rowlen = self.tableWidget.rowCount()
            collen = self.tableWidget.columnCount()

            with open(filename, "w") as csvfile:
                fieldnames = ['Rulename', 'Filename']
                writer = csv.DictWriter(csvfile, fieldnames=fieldnames)
                writer.writeheader()
                for row in range(rowlen):
                    rulename_cell = self.tableWidget.item(row, 0).text()
                    filename_cell = self.tableWidget.item(row, 1).text()
                    #print (row, rulename_cell, filename_cell)
                    writer.writerow({'Rulename':rulename_cell, 'Filename':filename_cell})

            os.system("open "+filename)
        except:
            return

    def loadRule(self):
        try:
            filename = QtWidgets.QFileDialog.getOpenFileName(None, "Load Yararule", self.configuration['lastRulePath'], "Text files (*.yar *.yara)")[0]
            config.set('YARAGROUND','lastRulePath', filename)
            with open(conf_filename,'w') as configfile:
                config.write(configfile)
            fp = open(filename, 'r')
            content = fp.read()
            self.ruleEditorBox.setPlainText(content)
            self.yararulePath.setText(filename)

        except:
            return

    def targetDIR(self):
        try:
            dirname = QtWidgets.QFileDialog.getExistingDirectory(None, '', self.configuration['lastTargetDIR'], QtWidgets.QFileDialog.ShowDirsOnly)
            config.set('YARAGROUND','lastTargetDIR', dirname)
            with open(conf_filename,'w') as configfile:
                config.write(configfile)
            print ("Target Directory :", dirname)
            self.dirPathEditor.setText(dirname)
        except:
            return
        
    def string_replace(self, content, targetdir):
        return content.replace(targetdir, "")

    def scan(self):
        self.tableWidget.setRowCount(0)
        resultfile = "result.txt"
        targetdir = self.dirPathEditor.text()
        recur_option = ""

        if self.recursiveCheckBox.isChecked() == True:
            recur_option = "-r"       
        command = "yara %s %s %s > %s" % (self.yararulePath.text(), targetdir, recur_option, resultfile)
        print (command)
        os.system(command)

        content = open(resultfile, 'r')
        #os.remove(resultfile)

        for item in content:
            item = self.string_replace(item, targetdir)
            rulename = item.split(" ")[0]
            filename = item.split(" ")[1]
            num_row = self.tableWidget.rowCount()
            self.tableWidget.insertRow(num_row)
            self.tableWidget.setItem(num_row, 0, QtWidgets.QTableWidgetItem(rulename))
            self.tableWidget.setItem(num_row, 1, QtWidgets.QTableWidgetItem(filename))

        msgBox = QtWidgets.QMessageBox(None)
        msgBox.setText("Yara Scan is Done!")
        msgBox.exec_()

    def eventFilter(self, object, event):
        print ("eventFilter")

    def keyPressEvent(self, ev):
        print ("keyPressEvent")
        if (ev.key() == Qt.Key_C) and (ev.modifiers() & Qt.ControlModifier): 
            self.copySelection()

    def copySelection(self):
        print ("copySelection")
        selection = self.tableWidget.selectedIndexes()
        if selection:
            rows = sorted(index.row() for index in selection)
            columns = sorted(index.column() for index in selection)
            rowcount = rows[-1] - rows[0] + 1
            colcount = columns[-1] - columns[0] + 1
            table = [[''] * colcount for _ in range(rowcount)]
            for index in selection:
                row = index.row() - rows[0]
                column = index.column() - columns[0]
                table[row][column] = index.data()
            stream = io.StringIO()
            csv.writer(stream).writerows(table)
            QApplication.clipboard().setText(stream.getvalue())

if __name__ == "__main__":
    import sys
    app = QtWidgets.QApplication(sys.argv)
    app.setStyleSheet(qdarkgraystyle.load_stylesheet())
    
    yaraground = QtWidgets.QDialog()
    ui = Ui_yaraground()
    ui.setupUi(yaraground)
    yaraground.show()
    sys.exit(app.exec_())
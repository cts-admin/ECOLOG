# -*- coding: utf-8 -*-
#================================================================================#
#  ECOLOG - Sistema Gerenciador de Banco de Dados para Levantamentos Ecológicos  #
#        ECOLOG - Database Management System for Ecological Surveys              #
#      Copyright (c) 1990-2016 Mauro J. Cavalcanti. All rights reserved.         #
#                                                                                #
#   Este programa é software livre; você pode redistribuí-lo e/ou                #
#   modificá-lo sob os termos da Licença Pública Geral GNU, conforme             #
#   publicada pela Free Software Foundation; tanto a versão 2 da                 #
#   Licença como (a seu critério) qualquer versão mais nova.                     #
#                                                                                # 
#   Este programa é distribuído na expectativa de ser útil, mas SEM              #
#   QUALQUER GARANTIA; sem mesmo a garantia implícita de                         #
#   COMERCIALIZAÇÃO ou de ADEQUAÇÃO A QUALQUER PROPÓSITO EM                      #
#   PARTICULAR. Consulte a Licença Pública Geral GNU para obter mais             #
#   detalhes.                                                                    #
#                                                                                #
#   This program is free software: you can redistribute it and/or                #
#   modify it under the terms of the GNU General Public License as published     #
#   by the Free Software Foundation, either version 2 of the License, or         #
#   version 3 of the License, or (at your option) any later version.             #
#                                                                                #
#   This program is distributed in the hope that it will be useful,              #
#   but WITHOUT ANY WARRANTY; without even the implied warranty of               #
#   MERCHANTABILITY or FITNESS FOR A PARTICULAR PURPOSE. See                     #
#   the GNU General Public License for more details.                             #
#                                                                                #
#  Dependências / Dependencies:                                                  #
#    Python 2.6+  (www.python.org)                                               #
#    PyQt 4.8+  (www.riverbankcomputing.com/software/pyqt)                       #
#================================================================================#

import os, sys
from PyQt4 import QtCore, QtGui

def rtrim(src, substr):
    s = src.mid(1, src.lastIndexOf(substr)-1)
    return s

class ChecklistDialog(QtGui.QDialog):
    def __init__(self, name, stringlist=None, checked=False, icon=None, parent=None):
        super(ChecklistDialog, self).__init__(parent)

        self.name = name
        self.icon = icon
        self.model = QtGui.QStandardItemModel()
        self.listView = QtGui.QListView()

        if stringlist is not None:
            for i in range(len(stringlist)):
                item = QtGui.QStandardItem(stringlist[i])
                item.setCheckable(True)
                check = QtCore.Qt.Checked if checked else QtCore.Qt.Unchecked
                item.setCheckState(check)
                self.model.appendRow(item)

        self.listView.setModel(self.model)
        layout = QtGui.QVBoxLayout()
        layout.addWidget(self.listView)

        buttonBox = QtGui.QDialogButtonBox(QtGui.QDialogButtonBox.Ok|QtGui.QDialogButtonBox.Cancel)
        layout.addWidget(buttonBox)
        self.setLayout(layout)
        self.setWindowTitle(self.name)
        if self.icon is not None: self.setWindowIcon(self.icon)

        self.connect(buttonBox, QtCore.SIGNAL("accepted()"), self, QtCore.SLOT("accept()"))
        self.connect(buttonBox, QtCore.SIGNAL("rejected()"), self, QtCore.SLOT("reject()"))

    def reject(self):
        QtGui.QDialog.reject(self)

    def accept(self):
        self.choices = []
        i = 0
        while self.model.item(i):
            if self.model.item(i).checkState():
                self.choices.append(self.model.item(i).text())
            i += 1
        QtGui.QDialog.accept(self)

class ChoiceDialog(QtGui.QDialog):
    def __init__(self, name, stringlist=None, multi=True, icon=None, parent=None):
        super(ChoiceDialog, self).__init__(parent)

        self.name = name
        self.icon = icon
        self.listWidget = QtGui.QListWidget()

        if multi:
            self.listWidget.setSelectionMode(QtGui.QListWidget.MultiSelection)
        else:
            self.listWidget.setSelectionMode(QtGui.QListWidget.SingleSelection)
            
        if stringlist is not None:
            self.listWidget.addItems(stringlist)
           
        layout = QtGui.QVBoxLayout()
        layout.addWidget(self.listWidget)
       
        buttonBox = QtGui.QDialogButtonBox(QtGui.QDialogButtonBox.Ok|QtGui.QDialogButtonBox.Cancel)
       
        layout.addWidget(buttonBox)
        self.setLayout(layout)
        self.setWindowTitle(self.name)
        if self.icon is not None: self.setWindowIcon(self.icon)

        self.connect(buttonBox, QtCore.SIGNAL("accepted()"), self, QtCore.SLOT("accept()"))
        self.connect(buttonBox, QtCore.SIGNAL("rejected()"), self, QtCore.SLOT("reject()"))
        self.listWidget.itemDoubleClicked.connect(self.itemDoubleClick)

    def reject(self):
        QtGui.QDialog.reject(self)

    def accept(self):
        self.choices = []
        for x in self.listWidget.selectedItems():
            self.choices.append(x.text())
        QtGui.QDialog.accept(self)
        
    def itemDoubleClick(self):
        self.accept()

class LoginDialog(QtGui.QDialog):
    def __init__(self, name, icon=None, parent=None):
        super(LoginDialog, self).__init__(parent)

        self.name = name
        self.icon = icon
        self.username = QtGui.QLineEdit()
        self.password = QtGui.QLineEdit()
        self.password.setEchoMode(QtGui.QLineEdit.Password)
        
        regexp = QtCore.QRegExp("\\S+")
        validator = QtGui.QRegExpValidator(regexp)
        self.username.setValidator(validator)
        self.password.setValidator(validator)
        
        loginLayout = QtGui.QFormLayout()
        loginLayout.addRow(self.trUtf8(u"Usuário:"), self.username)
        loginLayout.addRow(self.trUtf8("Senha:"), self.password)

        self.buttons = QtGui.QDialogButtonBox(QtGui.QDialogButtonBox.Ok|QtGui.QDialogButtonBox.Cancel)
        self.buttons.accepted.connect(self.accept)
        self.buttons.rejected.connect(self.reject)
        self.username.textChanged.connect(self.check_state)
        self.username.textChanged.emit(self.username.text())
        self.password.textChanged.connect(self.check_state)
        self.password.textChanged.emit(self.username.text())

        layout = QtGui.QVBoxLayout()
        layout.addLayout(loginLayout)
        layout.addWidget(self.buttons)
        self.setLayout(layout)
        
        self.setWindowTitle(self.name)
        if self.icon is not None: self.setWindowIcon(self.icon)
        
    def accept(self):
        if len(self.username.text()) > 0 and len(self.password.text()) > 0:
            QtGui.QDialog.accept(self)
            
    def check_state(self, *args, **kwargs):
        sender = self.sender()
        validator = sender.validator()
        state = validator.validate(sender.text(), 0)[0]
        if state == QtGui.QValidator.Acceptable:
            color = "#c4df9b" # green
        elif state == QtGui.QValidator.Intermediate:
            color = "#fff79a" # yellow
        else:
            color = "#f6989d" # red
        sender.setStyleSheet("QLineEdit { background-color: %s }" % color)
        
    def keyPressEvent(self, event):
        if event.key() == QtCore.Qt.Key_Return:
            if self.username.hasFocus() and len(self.username.text()) > 0:
                self.password.setFocus()
                event.ignore()
            if self.password.hasFocus() and len(self.password.text()) > 0:
                event.accept()
                self.accept()

class QueryDialog(QtGui.QDialog):
    def __init__(self, title, tableName, fieldList, fieldTypes, displayNames=None, 
                alias=True, SQLSelectClause=False, icon=None, parent=None):
        super(QueryDialog, self).__init__(parent)
        
        self.title = title
        self.icon = icon
        self.tableName = tableName
        self.fieldList = fieldList
        self.fieldTypes = fieldTypes
        if displayNames is not None: 
            self.displayNames = displayNames
        else:
            self.displayNames = fieldList
        self.alias = alias
        self.SQLSelectClause = SQLSelectClause
        self.SQLWhereClause = []
        self.operM = [">", ">=", "<", "<=", "=", "<>", "LIKE", "NOT LIKE"]
        self.operL = ["AND", "OR"]
        self.SQL = ""
        
        tableLabel = QtGui.QLabel(self.trUtf8(u"Tabela:"))
        self.comboTable = QtGui.QComboBox()
        self.comboTable.addItems(self.tableName)
        fieldLabel = QtGui.QLabel(self.trUtf8(u"Campos:"))
        self.comboFields = QtGui.QComboBox()
        self.comboFields.addItems(self.displayNames)
        operatorLabel = QtGui.QLabel(self.trUtf8(u"Operador:"))
        self.comboConditional = QtGui.QComboBox()
        self.comboConditional.addItem(self.trUtf8(u"Maior"))
        self.comboConditional.addItem(self.trUtf8(u"Maior ou Igual"))
        self.comboConditional.addItem(self.trUtf8(u"Menor"))
        self.comboConditional.addItem(self.trUtf8(u"Menor ou Igual"))
        self.comboConditional.addItem(self.trUtf8(u"Igual"))
        self.comboConditional.addItem(self.trUtf8(u"Diferente"))
        self.comboConditional.addItem(self.trUtf8(u"Contendo"))
        self.comboConditional.addItem(self.trUtf8(u"Não Contendo"))
        valueLabel = QtGui.QLabel(self.trUtf8(u"Valor:"))
        self.editFilter = QtGui.QComboBox()
        self.editFilter.setEditable(True)
        self.rbAND = QtGui.QRadioButton(self.trUtf8(u"E"))
        self.rbOR = QtGui.QRadioButton(self.trUtf8(u"Ou"))
        self.IncludeBtn = QtGui.QPushButton()
        self.IncludeBtn.setIcon(QtGui.QIcon(":/plus.png"))
        self.IncludeBtn.setToolTip(self.trUtf8(u"Incluir condição"))
        filtersLabel = QtGui.QLabel(self.trUtf8(u"Lista de Condições:"))
        self.listFilters = QtGui.QListWidget()
        self.ClearBtn = QtGui.QPushButton()
        self.ClearBtn.setIcon(QtGui.QIcon(":/clear.png"))
        self.ClearBtn.setToolTip(self.trUtf8(u"Limpar lista de condições"))    
        self.ExcludeBtn = QtGui.QPushButton()
        self.ExcludeBtn.setIcon(QtGui.QIcon(":/minus.png"))
        self.ExcludeBtn.setToolTip(self.trUtf8(u"Excluir condição"))
        buttonBox = QtGui.QDialogButtonBox(QtGui.QDialogButtonBox.Ok|QtGui.QDialogButtonBox.Cancel)
                                            
        topLayout = QtGui.QGridLayout()
        topLayout.addWidget(tableLabel, 0, 0)
        topLayout.addWidget(fieldLabel, 0, 1)
        topLayout.addWidget(operatorLabel, 0, 2)
        topLayout.addWidget(self.comboTable, 1, 0)
        topLayout.addWidget(self.comboFields, 1, 1)
        topLayout.addWidget(self.comboConditional, 1, 2)
        topLayout.addWidget(valueLabel, 2, 0)
        topLayout.addWidget(self.editFilter, 3, 0, 1, -1)
        midLayout = QtGui.QHBoxLayout()
        midLayout.addWidget(self.rbAND)
        midLayout.addWidget(self.rbOR)
        midLayout.addStretch(1)
        midLayout.addWidget(self.IncludeBtn)
        btmLayout = QtGui.QVBoxLayout()
        btmLayout.addWidget(filtersLabel)
        btmLayout.addWidget(self.listFilters)
        btnLayout = QtGui.QHBoxLayout()
        btnLayout.addStretch(1)
        btnLayout.addWidget(self.ClearBtn)
        btnLayout.addWidget(self.ExcludeBtn)
        mainLayout = QtGui.QVBoxLayout()
        mainLayout.addLayout(topLayout)
        mainLayout.addLayout(midLayout)
        mainLayout.addLayout(btmLayout)
        mainLayout.addLayout(btnLayout)
        mainLayout.addWidget(buttonBox)
        self.setLayout(mainLayout)
        
        self.connect(buttonBox, QtCore.SIGNAL("accepted()"), self.accept)
        self.connect(buttonBox, QtCore.SIGNAL("rejected()"), self.reject)
        self.comboFields.currentIndexChanged["QString"].connect(self.comboFieldsChange)
        self.IncludeBtn.clicked.connect(self.IncludeBtnClick)
        self.ClearBtn.clicked.connect(self.ClearBtnClick)
        self.ExcludeBtn.clicked.connect(self.ExcludeBtnClick)
        
        self.setWindowTitle(self.title)
        if self.icon is not None: self.setWindowIcon(QtGui.QIcon(self.icon))
        self.EnableControls()
            
    def comboFieldsChange(self):
        self.editFilter.lineEdit().setText("")
    
    def IncludeBtnClick(self):
        if not self.editFilter.lineEdit().text(): return
        self.editFilter.addItem(self.editFilter.lineEdit().text())
        i = self.comboFields.currentIndex()
        j = self.comboConditional.currentIndex()
        k = self.comboTable.currentIndex()
        s = self.fieldList[i]
        if self.fieldTypes[i] == "text":
            openQuote = ' "'
            closeQuote = '"'
        elif self.fieldTypes[i] == "numeric":
            openQuote = ' '
            closeQuote = ''
        if self.alias:
            selection = self.comboTable.currentText() + '."' + self.displayNames[i] + '" ' + \
                self.comboConditional.currentText() + openQuote +  self.editFilter.lineEdit().text() + closeQuote    
        else:
            selection = self.displayNames[i] + " " + \
            self.comboConditional.currentText() + openQuote +  self.editFilter.lineEdit().text() + closeQuote
        if not self.listFilters.findItems(selection, QtCore.Qt.MatchExactly):
            self.listFilters.addItem(selection)
            if s.find('.') == -1:
                if self.alias:
                    condition = self.tableName[k] + '.' + s + ' ' + self.operM[j]
                else:
                    condition = s + ' ' + self.operM[j]
            else:
                condition = s + ' ' + self.operM[j]
            if j > 5:
                value = ' "%' + self.editFilter.lineEdit().text() + '%"'
            else:
                value = openQuote + self.editFilter.lineEdit().text() + closeQuote
            SQLStr = condition + value
            self.SQLWhereClause.append(SQLStr)
            self.EnableControls()
    
    def ClearBtnClick(self):
        self.SQL = ""
        self.SQLWhereClause = []
        self.comboFields.setCurrentIndex(0)
        self.comboConditional.setCurrentIndex(0)
        self.editFilter.lineEdit().setText("")
        self.listFilters.clear()
        self.EnableControls()
    
    def ExcludeBtnClick(self):
        if self.listFilters.count() > 0:
            index = self.listFilters.currentRow()
            if index != -1: 
                item = self.listFilters.takeItem(index)
                del item
                self.SQLWhereClause.pop(index)
            self.EnableControls()
            
    def reject(self):
        QtGui.QDialog.reject(self)

    def accept(self):
        if self.SQLSelectClause:
            self.SQL = "SELECT * FROM " + self.tableName[self.comboTable.currentIndex()]
        else:
            self.SQL = ""
        if len(self.SQLWhereClause) > 0:
            self.SQL = self.SQL + " WHERE "
            for i in range(len(self.SQLWhereClause)):
                self.SQL = self.SQL + self.SQLWhereClause[i]
                if self.rbAND.isChecked():
                    self.SQL = self.SQL + " AND "
                elif self.rbOR.isChecked():
                    self.SQL = self.SQL + " OR "
            if self.rbAND.isChecked():		
                self.SQL = rtrim(self.SQL, " AND ")
            elif self.rbOR.isChecked():
                self.SQL = rtrim(self.SQL, " OR ")
        self.SQL = QtCore.QString(self.SQL).toUtf8()
        QtGui.QDialog.accept(self)
        
    def EnableControls(self):
        self.comboTable.setEnabled(False)
        self.editFilter.lineEdit().clear()
        self.rbAND.setEnabled(self.listFilters.count() > 0)
        self.rbOR.setEnabled(self.listFilters.count() > 0)
        if self.rbAND.isEnabled():
            self.rbAND.setChecked(self.rbAND.isEnabled())
            self.rbOR.setChecked(not self.rbAND.isEnabled())
        else:
            self.rbAND.setChecked(self.rbAND.isEnabled())
            self.rbOR.setChecked(self.rbAND.isEnabled())
        self.ExcludeBtn.setEnabled(self.listFilters.count() > 0)
        self.ClearBtn.setEnabled(self.listFilters.count() > 0)
        
    def keyPressEvent(self, event):
        if event.key() == QtCore.Qt.Key_Return:
            self.IncludeBtnClick()
            event.ignore()
#!/usr/bin/python
# -*- coding: utf-8 -*-
# ============================================================================== #
#                               ECOLOG                                           #
# ============================================================================== #
#                                                                                #
# Original Attribution:                                                          #
#   Sistema Gerenciador de Banco de Dados para Levantamentos Ecológicos          #
#   ECOLOG - Database Management System for Ecological Surveys                   #
#   Mauro J. Cavalcanti.                                                         #
#                                                                                #
# ****************************************************************************** #
# This copy of the software is now maintained by Conservation Technology         #
# Solution Inc.                                                                  #
# ****************************************************************************** #
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
#    Python 3.6+ (www.python.org)                                                #
#    PyQt  5.10.1+ (www.riverbankcomputing.com/software/pyqt)                    #
#    PyYAML 4.1+ (www.pyyaml.org)                                                #
#                                                                                #
# ============================================================================== #

import os
import platform
import sqlite3
import sys
import time
import yaml
from datetime import datetime

from PyQt5 import QtCore, QtGui
from PyQt5.QtCore import QT_VERSION_STR
from os.path import basename

import Export
import Import
from Dialog import ChoiceDialog, LoginDialog, QueryDialog
from Report import Report
from Sheet import Sheet
from Useful import iif

__version__ = "5.1"


class MainWindow(QtGui.QMainWindow):
    def __init__(self):
        super(MainWindow, self).__init__()
        self.projectfile = ""
        self.projectdata = {}
        self.mdi = QtGui.QMdiArea()
        self.setCentralWidget(self.mdi)

        self.mdi.subWindowActivated.connect(self.update_menus)
        self.windowMapper = QtCore.QSignalMapper(self)
        self.windowMapper.mapped[QtGui.QWidget].connect(self.mdi.setActiveSubWindow)

        self.create_actions()
        self.create_menus()
        self.create_toolbars()
        self.create_statusbar()
        self.update_menus()

        self.restore_settings()
        self.setUnifiedTitleAndToolBarOnMac(True)

    def close_event(self, event):
        reply = QtGui.QMessageBox.question(self, self.trUtf8(u"Confirmação"),
                                           self.trUtf8(u"Deseja encerrar o programa?"),
                                           QtGui.QMessageBox.Yes, QtGui.QMessageBox.No)
        if reply == QtGui.QMessageBox.Yes:
            event.accept()
            self.save_settings()
            self.mdi.closeAllSubWindows()
        else:
            event.ignore()

    def file_new(self):
        filename = QtGui.QFileDialog.getSaveFileName(self,
                                                     self.trUtf8(u"Salvar Como"), os.getcwd(),
                                                     self.trUtf8(u"Metadados (*.yml)"))
        if not filename.isEmpty():
            if any(self.projectdata):
                self.file_close()
            os.chdir(str(QtCore.QFileInfo(filename).path()))
            self.projectfile = str(filename)
            self.init_project_data()
            result = fedit(self.datagroup,
                           title=self.trUtf8(u"Informações do Projeto"),
                           icon=QtGui.QIcon(":/ecolog.png"),
                           parent=self)
            if result is not None:
                self.get_project_data(result)
                with open(self.projectfile, 'w') as outfile:
                    yaml.dump(self.projectdata, outfile, default_flow_style=False)
                self.update_menus()
                self.fileAddAction.setEnabled(True)
                self.fileEditAction.setEnabled(True)
                self.fileCloseAction.setEnabled(True)
                self.fileImportAction.setEnabled(True)
                reply = QtGui.QMessageBox.question(self, self.trUtf8(u"Confirmação"),
                                                   self.trUtf8(u"Deseja adicionar planilhas ao projeto?"),
                                                   QtGui.QMessageBox.Yes, QtGui.QMessageBox.No)
                if reply == QtGui.QMessageBox.Yes:
                    self.file_add()

    def file_open(self):
        filename = QtGui.QFileDialog.getOpenFileName(self,
                                                     self.trUtf8(u"Abrir"), os.getcwd(),
                                                     self.trUtf8(u"Projetos (*.yml)"))
        if not filename.isEmpty():
            if any(self.projectdata):
                self.file_close()
            os.chdir(str(QtCore.QFileInfo(filename).path()))
            self.projectfile = str(filename)
            with open(self.projectfile, 'r') as infile:
                self.projectdata = yaml.load(infile)
            ok = False
            if self.projectdata.has_key("dataset1") and len(self.projectdata["dataset1"]) > 0:
                ok = True
                fname1 = self.projectdata["dataset1"]
                self.load_file(fname1)
            if self.projectdata.has_key("dataset2")  and len(self.projectdata["dataset2"]) > 0:
                ok = True
                fname2 = self.projectdata["dataset2"]
                self.load_file(fname2)
            if ok:
                self.mdi.activatePreviousSubWindow()
                self.mdi.cascadeSubWindows()
                self.update_menus()
            if not ok:
                QtGui.QMessageBox.warning(self, self.trUtf8(u"Aviso"),
                                          self.trUtf8(u"O projeto não contém arquivos de dados"))
                self.fileAddAction.setEnabled(True)
                self.fileEditAction.setEnabled(True)
                self.fileCloseAction.setEnabled(True)
                self.fileImportAction.setEnabled(True)

    def file_edit(self):
        self.set_project_data()
        result = fedit(self.datagroup,
                       title=self.trUtf8(u"Informações do Projeto"),
                       icon=QtGui.QIcon(":/ecolog.png"),
                       parent=self)
        if result is not None:
            self.get_project_data(result)
            with open(self.projectfile, 'w') as outfile:
                yaml.dump(self.projectdata, outfile, default_flow_style=False)

    def file_reload(self):
        self.mdi.closeAllSubWindows()
        fname1 = self.projectdata["dataset1"]
        self.load_file(fname1)
        if self.projectdata.has_key("dataset2") and len(self.projectdata["dataset2"]) > 0:
            fname2 = self.projectdata["dataset2"]
            self.load_file(fname2)
        self.mdi.activatePreviousSubWindow()
        self.mdi.cascadeSubWindows()
        self.update_menus()

    def file_close(self):
        self.projectfile = ""
        self.projectdata = {}
        self.mdi.closeAllSubWindows()
        self.update_menus()

    def file_add(self):
        options = [(self.trUtf8(u"Planilha:"),
                    [0, self.trUtf8(u"Planilha de Dados"),
                     self.trUtf8(u"Planilha de Variáveis")])]
        result = fedit(options,
                       title=self.trUtf8(u"Adicionar"),
                       icon=QtGui.QIcon(":/ecolog.png"),
                       parent=self)
        if result is not None:
            index = result[0]
            filename = QtGui.QFileDialog.getOpenFileName(self,
                                                         self.trUtf8(u"Abrir"), os.getcwd(),
                                                         self.trUtf8(u"Planilhas (*.csv *.docs *.ods *.tsv *.xls *.xlsx)"))
            if not filename.isEmpty():
                if index == 0:
                    self.projectdata["dataset1"] = basename(str(filename))
                    self.load_file(str(filename))
                else:
                    self.projectdata["dataset2"] = basename(str(filename))
                    self.load_file(str(filename))
                with open(self.projectfile, 'w') as outfile:
                    yaml.dump(self.projectdata, outfile, default_flow_style=False)
                self.file_reload()

    def file_remove(self):
        options = [(self.trUtf8(u"Planilha:"),
                    [0, self.trUtf8(u"Planilha de Dados"),
                     self.trUtf8(u"Planilha de Variáveis")])]
        result = fedit(options,
                       title=self.trUtf8(u"Remover"),
                       icon=QtGui.QIcon(":/ecolog.png"),
                       parent=self)
        if result is not None:
            index = result[0]
            if index == 0:
                for window in self.mdi.subWindowList():
                    if window.windowTitle() == basename(self.projectdata["dataset1"]):
                        window.close()
                self.projectdata["dataset1"] = ""
            else:
                for window in self.mdi.subWindowList():
                    if window.windowTitle() == basename(self.projectdata["dataset2"]):
                        window.close()
                self.projectdata["dataset2"] = ""
            with open(self.projectfile, 'w') as outfile:
                yaml.dump(self.projectdata, outfile)

    def file_import(self):
        msgBox = QtGui.QMessageBox()
        msgBox.setWindowTitle(self.trUtf8(u"Importar"))
        msgBox.setIcon(QtGui.QMessageBox.Question)
        msgBox.setText(self.trUtf8(u"Selecione o destino do novo projeto\n(O projeto atual deverá estar vazio)"))
        btnQS = QtGui.QPushButton(self.trUtf8(u" Adicionar ao projeto atual "))
        msgBox.addButton(btnQS, QtGui.QMessageBox.YesRole)
        btnNo = QtGui.QPushButton(self.trUtf8(u" Criar novo projeto "))
        msgBox.addButton(btnNo, QtGui.QMessageBox.NoRole)
        btnCancel = QtGui.QPushButton(self.trUtf8(u" Cancelar "))
        msgBox.addButton(btnCancel, QtGui.QMessageBox.RejectRole)
        ret = msgBox.exec_()
        if ret == 0:
            if self.projectdata.has_key("dataset1") and len(self.projectdata["dataset1"]) > 0:
                QtGui.QMessageBox.warning(self, self.trUtf8(u"Aviso"),
                                          self.trUtf8(u"O projeto já contém um arquivo de dados"))
                return

        elif ret == 1:
            filename = QtGui.QFileDialog.getSaveFileName(self,
                                                         self.trUtf8(u"Salvar Como"), os.getcwd(),
                                                         self.trUtf8(u"Metadados (*.yml)"))
            if not filename.isEmpty():
                if any(self.projectdata):
                    self.file_close()
                os.chdir(str(QtCore.QFileInfo(filename).path()))
                self.projectfile = str(filename)
                self.init_project_data()
                result = fedit(self.datagroup,
                               title=self.trUtf8(u"Informações do Projeto"),
                               icon=QtGui.QIcon(":/ecolog.png"),
                               parent=self)
                if result is not None:
                    self.get_project_data(result)
                    with open(self.projectfile, 'w') as outfile:
                        yaml.dump(self.projectdata, outfile, default_flow_style=False)
                    self.update_menus()
                    self.fileAddAction.setEnabled(True)
                    self.fileEditAction.setEnabled(True)
                    self.fileCloseAction.setEnabled(True)
                    self.fileImportAction.setEnabled(True)
            else:
                return

        elif ret == 2:
            return

        formats = ["csv", "xls", "xlsx", "ods"]
        options = [(self.trUtf8(u"Formato de Entrada:"),
                    [0, self.trUtf8(u"Planilha GBIF"), self.trUtf8(u"Planilha OBIS"),
                     self.trUtf8(u"Planilha TEAM"), self.trUtf8(u"Planilha VertNet")]),
                   (self.trUtf8(u"Formato de Saída"),
                    [0, self.trUtf8(u"Texto CSV (.csv)"),
                     self.trUtf8(u"Microsoft Excel 97/2000/XP/2003 (.xls)"),
                     self.trUtf8(u"Microsoft Excel 2007/2010 (.xlsx)"),
                     self.trUtf8(u"Planilha ODF (.ods)")])]
        result = fedit(options,
                       title=self.trUtf8(u"Importar"),
                       icon=QtGui.QIcon(":/ecolog.png"),
                       parent=self)
        if result is not None:
            choice = result[0]
            format = formats[result[1]]
            if choice == 2:
                protocols = ["Avian","Butterfly","Climate","Primate","Tree","Liana"]
                choices = [(self.trUtf8(u"Protocolo:"),
                            [0,self.trUtf8(u"Aves"), self.trUtf8(u"Borboletas"), self.trUtf8(u"Clima"),
                             self.trUtf8(u"Primatas"), self.trUtf8(u"Vegetação - Árvores"),
                             self.trUtf8(u"Vegetação - Lianas")])]
                result = fedit(choices,
                               title=self.trUtf8(u"Selecionar"),
                               icon=QtGui.QIcon(":/ecolog.png"),
                               parent=self)
                if result is not None:
                    protocol = protocols[result[0]]
                else:
                    return
            filename = QtGui.QFileDialog.getOpenFileName(self,
                                                         self.trUtf8(u"Abrir"), os.getcwd(),
                                                         self.trUtf8(u"Dados (*.csv *.txt *.zip)"))
            if not filename.isEmpty():
                if choice == 0:
                    recs, infile = Import.fromGBIF(filename, format)

                elif choice == 1:
                    recs, infile = Import.fromOBIS(filename, format)

                elif choice == 2:
                    recs, infile = Import.fromTEAM(filename, format, protocol)

                elif choice == 3:
                    recs, infile = Import.fromVertNet(filename, format)

                QtGui.QMessageBox.information(self, self.trUtf8(u"Informação"),
                                              self.trUtf8(u"Dados importados para:\n") + \
                                              str(recs) + self.trUtf8(u" registros"))

                self.projectdata["dataset1"] = basename(infile)
                self.load_file(infile)

    def file_export(self):
        dataset = self.projectdata["dataset1"]
        window = self.find_mdi_child(dataset)
        if window:
            self.mdi.setActiveSubWindow(window)
            sheet = self.active_mdi_child()
            options = [(self.trUtf8(u"Formato:"),
                        [0, self.trUtf8(u"Metadados em EML"),
                         self.trUtf8(u"Arquivos do Fitopac 1"), self.trUtf8(u"Arquivo do Fitopac 2"),
                         self.trUtf8(u"Coordenadas em KML"), self.trUtf8(u"Shapefile ESRI"),
                         self.trUtf8(u"Arquivo RDE/BRAHMS"), self.trUtf8(u"Matriz CEP"),
                         self.trUtf8(u"Matriz MVSP"), self.trUtf8(u"Matriz Fitopac"),
                         self.trUtf8(u"Matriz CSV")])]
            result = fedit(options,
                           title=self.trUtf8(u"Exportar"),
                           icon=QtGui.QIcon(":/ecolog.png"),
                           parent=self)
            if result is not None:
                format = result[0]

                if format == 0:
                    filename = QtGui.QFileDialog.getSaveFileName(self,
                                                                 self.trUtf8(u"Salvar Como"), os.getcwd(),
                                                                 self.trUtf8(u"Ecological Metadata Language (*.eml)"))
                    if not filename.isEmpty():
                        levels = Export.toEML(filename=filename,
                                              projectdata=self.projectdata,
                                              headers=sheet.headers,
                                              types=sheet.types,
                                              data=sheet.data)
                        QtGui.QMessageBox.information(self, self.trUtf8(u"Informação"),
                                                      self.trUtf8(u"Metadados gravados no arquivo:\n") + os.path.basename(str(filename)))

                elif format == 1:
                    if self.projectdata["method"].toLower() == self.trUtf8(u"parcela"):
                        kind = 1
                    elif self.projectdata["method"].toLower() == self.trUtf8(u"quadrante"):
                        kind = 2
                    else:
                        QtGui.QMessageBox.warning(self, self.trUtf8(u"Informação"),
                                                  self.trUtf8(u"Tipo de levantamento deve ser parcela ou quadrante"))
                        return
                    filename = QtGui.QFileDialog.getSaveFileName(self,
                                                                 self.trUtf8(u"Salvar Como"), os.getcwd(),
                                                                 self.trUtf8(u"Arquivos Fitopac1 (*.dad)"))
                    if not filename.isEmpty():
                        nfam, nesp, nind = Export.toFitopac1(filename=filename,
                                                             projectdata=self.projectdata,
                                                             headers=sheet.headers,
                                                             types=sheet.types,
                                                             data=sheet.data,
                                                             tipo=kind)
                        QtGui.QMessageBox.information(self, self.trUtf8(u"Informação"),
                                                      self.trUtf8(u"Dados gravados no arquivo:\n") + os.path.basename(str(filename)) + \
                                                      self.trUtf8(u" para:\n") + str(nfam) + self.trUtf8(u" famílias \n") + \
                                                      str(nesp) + self.trUtf8(u" espécies\n") + str(nind) + self.trUtf8(u" indivíduos"))

                elif format == 2:
                    filename = QtGui.QFileDialog.getSaveFileName(self,
                                                                 self.trUtf8(u"Salvar Como"), os.getcwd(),
                                                                 self.trUtf8(u"Arquivos Fitopac2 (*.fpd)"))
                    if not filename.isEmpty():
                        nfam, nesp, nind, nsu = Export.toFitopac2(filename=filename,
                                                                  projectdata=self.projectdata,
                                                                  headers=sheet.headers,
                                                                  types=sheet.types,
                                                                  data=sheet.data)
                        QtGui.QMessageBox.information(self, self.trUtf8(u"Informação"),
                                                      self.trUtf8(u"Dados gravados no arquivo:\n") + os.path.basename(str(filename)) + \
                                                      self.trUtf8(u" para:\n") + str(nfam) + self.trUtf8(u" famílias \n") + \
                                                      str(nesp) + self.trUtf8(u" espécies\n") + str(nind) + self.trUtf8(u" indivíduos\n") + \
                                                      str(nsu) + self.trUtf8(u" amostras"))

                elif format == 3:
                    filename = QtGui.QFileDialog.getSaveFileName(self,
                                                                 self.trUtf8(u"Salvar Como"), os.getcwd(),
                                                                 self.trUtf8(u"Keyhole Markup Language (*.kml)"))
                    if not filename.isEmpty():
                        recs = Export.toKML(filename=filename,
                                            projectdata=self.projectdata,
                                            headers=sheet.headers,
                                            types=sheet.types,
                                            data=sheet.data)
                        QtGui.QMessageBox.information(self, self.trUtf8(u"Informação"),
                                                      str(recs) + self.trUtf8(u" registros gravados no arquivo:\n") + \
                                                      os.path.basename(str(filename)))

                elif format == 4:
                    filename = QtGui.QFileDialog.getSaveFileName(self,
                                                                 self.trUtf8(u"Salvar Como"), os.getcwd(),
                                                                 self.trUtf8(u"Shapefiles (*.shp)"))
                    if not filename.isEmpty():
                        recs = Export.toSHP(filename=filename,
                                            projectdata=self.projectdata,
                                            headers=sheet.headers,
                                            types=sheet.types,
                                            data=sheet.data)
                        QtGui.QMessageBox.information(self, self.trUtf8(u"Informação"),
                                                      str(recs) + self.trUtf8(u" registros gravados no arquivo:\n") + \
                                                      os.path.basename(str(filename)))

                elif format == 5:
                    filename = QtGui.QFileDialog.getSaveFileName(self,
                                                                 self.trUtf8(u"Salvar Como"), os.getcwd(),
                                                                 self.trUtf8(u"dBase (*.dbf)"))
                    if not filename.isEmpty():
                        recs = Export.toRDE(filename=filename,
                                            projectdata=self.projectdata,
                                            headers=sheet.headers,
                                            types=sheet.types,
                                            data=sheet.data)
                        QtGui.QMessageBox.information(self, self.trUtf8(u"Informação"),
                                                      str(recs) + self.trUtf8(u" registros gravados no arquivo:\n") + \
                                                      os.path.basename(str(filename)))

                elif format == 6:
                    filename = QtGui.QFileDialog.getSaveFileName(self,
                                                                 self.trUtf8(u"Salvar Como"), os.getcwd(),
                                                                 self.trUtf8(u"Cornell Ecology Programs (*.dta)"))
                    if not filename.isEmpty():
                        nrows, ncols = Export.toCEP(filename=filename,
                                                    projectdata=self.projectdata,
                                                    headers=sheet.headers,
                                                    types=sheet.types,
                                                    data=sheet.data)
                    QtGui.QMessageBox.information(self, self.trUtf8(u"Informação"),
                                                  self.trUtf8(u"Dados gravados no arquivo:\n") + os.path.basename(str(filename)) + \
                                                  self.trUtf8(u" para:\n") + str(nrows) + self.trUtf8(u" amostras \n") + \
                                                  str(ncols) + self.trUtf8(u" espécies"))

                elif format == 7:
                    choices = [(self.trUtf8(u"Tipo:"),
                                [1, self.trUtf8(u"Presença/Ausência"), self.trUtf8(u"Número de Indivíduos")])]
                    result = fedit(choices,
                                   title=self.trUtf8(u"Matriz de Dados"),
                                   icon=QtGui.QIcon(":/ecolog.png"),
                                   parent=self)
                    if result is not None:
                        kind = result[0] + 1
                        filename = QtGui.QFileDialog.getSaveFileName(self,
                                                                     self.trUtf8(u"Salvar Como"), os.getcwd(),
                                                                     self.trUtf8(u"Multivariate Statistical Package (*.mvs)"))
                        if not filename.isEmpty():
                            nrows, ncols = Export.toMatrix(filename=filename,
                                                           projectdata=self.projectdata,
                                                           headers=sheet.headers,
                                                           types=sheet.types,
                                                           data=sheet.data,
                                                           format="mvsp",
                                                           kind=kind)
                            QtGui.QMessageBox.information(self, self.trUtf8(u"Informação"),
                                                          self.trUtf8(u"Dados gravados no arquivo:\n") + os.path.basename(str(filename)) + \
                                                          self.trUtf8(u" para:\n") + str(nrows) + self.trUtf8(u" amostras \n") + \
                                                          str(ncols) + self.trUtf8(u" espécies"))

                elif format == 8:
                    choices = [(self.trUtf8(u"Tipo:"),
                                [1, self.trUtf8(u"Presença/Ausência"), self.trUtf8(u"Número de Indivíduos")])]
                    result = fedit(choices,
                                   title=self.trUtf8(u"Matriz de Dados"),
                                   icon=QtGui.QIcon(":/ecolog.png"),
                                   parent=self)
                    if result is not None:
                        kind = result[0] + 1
                        filename = QtGui.QFileDialog.getSaveFileName(self,
                                                                     self.trUtf8(u"Salvar Como"), os.getcwd(),
                                                                     self.trUtf8(u"Matriz Fitopac (*.fpm)"))
                        if not filename.isEmpty():
                            nrows, ncols = Export.toFPM(filename=filename,
                                                        projectdata=self.projectdata,
                                                        headers=sheet.headers,
                                                        types=sheet.types,
                                                        data=sheet.data,
                                                        kind=kind)
                            QtGui.QMessageBox.information(self, self.trUtf8(u"Informação"),
                                                          self.trUtf8(u"Dados gravados no arquivo:\n") + os.path.basename(str(filename)) + \
                                                          self.trUtf8(u" para:\n") + str(nrows) + self.trUtf8(u" amostras \n") + \
                                                          str(ncols) + self.trUtf8(u" espécies"))

                elif format == 9:
                    choices = [(self.trUtf8(u"Tipo:"),
                                [1, self.trUtf8(u"Presença/Ausência"), self.trUtf8(u"Número de Indivíduos")])]
                    result = fedit(choices,
                                   title=self.trUtf8(u"Matriz de Dados"),
                                   icon=QtGui.QIcon(":/ecolog.png"),
                                   parent=self)
                    if result is not None:
                        kind = result[0] + 1
                        filename = QtGui.QFileDialog.getSaveFileName(self,
                                                                     self.trUtf8(u"Salvar Como"), os.getcwd(),
                                                                     self.trUtf8(u"CSV (separado por vírgulas) (*.csv)"))
                        if not filename.isEmpty():
                            nrows, ncols = Export.toCSV(filename=filename,
                                                        projectdata=self.projectdata,
                                                        headers=sheet.headers,
                                                        types=sheet.types,
                                                        data=sheet.data,
                                                        kind=kind)
                            QtGui.QMessageBox.information(self, self.trUtf8(u"Informação"),
                                                          self.trUtf8(u"Dados gravados no arquivo:\n") + os.path.basename(str(filename)) + \
                                                          self.trUtf8(u" para:\n") + str(nrows) + self.trUtf8(u" amostras \n") + \
                                                          str(ncols) + self.trUtf8(u" espécies"))

    def data_sort(self):
        sheet = self.active_mdi_child()
        col = sheet.currentColumn()
        sheet.sortItems(col)
        sheet.setCurrentCell(0, col)

    def data_find(self):
        self.reset_search()
        searchStr, ok = QtGui.QInputDialog.getText(self, self.trUtf8(u"Pesquisar"),
                                                   self.trUtf8(u"Digite o texto a ser pesquisado:"))
        if ok:
            sheet = self.active_mdi_child()
            items = sheet.findItems(searchStr, QtCore.Qt.MatchContains)
            if items:
                for item in items:
                    sheet.setItemSelected(item, True)
                    sheet.item(item.row(), item.column()).setBackground(QtGui.QColor(0,255,51))
                QtGui.QMessageBox.information(self, self.trUtf8(u"Informação"),
                                              str(len(items)) + self.trUtf8(u" ocorrência(s) encontrada(s)"))
            else:
                QtGui.QMessageBox.warning(self, self.trUtf8(u"Aviso"),
                                          self.trUtf8(u"Texto não encontrado"))

    def data_filter(self):
        sheet = self.active_mdi_child()
        sheet.setData()
        filter = self.qbe(sheet, False)
        if len(filter) > 0:
            if sheet.setFilter(filter) == 0:
                QtGui.QMessageBox.warning(self, self.trUtf8(u"Aviso"),
                                          self.trUtf8(u"Não há registros nesta condição"))

    def report_catalog(self):
        dataset = self.projectdata["dataset1"]
        window = self.find_mdi_child(dataset)
        if window:
            self.mdi.setActiveSubWindow(window)
            sheet = self.active_mdi_child()
            if len(sheet.headers) < 12:
                QtGui.QMessageBox.warning(self, self.trUtf8(u"Aviso"),
                                          self.trUtf8(u"Não há campos suficientes na planilha"))
                return
            filter_expr = self.get_filter(sheet)
            filename = QtGui.QFileDialog.getSaveFileName(self,
                                                         self.trUtf8(u"Salvar Como"), os.getcwd(),
                                                         self.trUtf8(u"Relatórios (*.htm *.html)"))
            if not filename.isEmpty():
                report = Report(filename=filename,
                                projectdata=self.projectdata,
                                headers=sheet.headers,
                                types=sheet.types,
                                data=sheet.data,
                                filter=filter_expr)
                report.Checklist()
                self.mdi.addSubWindow(report)
                self.mdi.setWindowIcon(QtGui.QIcon(":/web.png"))
                report.show()

    def report_labels(self):
        dataset = self.projectdata["dataset1"]
        window = self.find_mdi_child(dataset)
        if window:
            self.mdi.setActiveSubWindow(window)
            sheet = self.active_mdi_child()
            if len(sheet.headers) < 12:
                QtGui.QMessageBox.warning(self, self.trUtf8(u"Aviso"),
                                          self.trUtf8(u"Não há campos suficientes na planilha"))
                return
            filter_expr = self.get_filter(sheet)
            filename = QtGui.QFileDialog.getSaveFileName(self,
                                                         self.trUtf8(u"Salvar Como"), os.getcwd(),
                                                         self.trUtf8(u"Relatórios (*.htm *.html)"))
            if not filename.isEmpty():
                options = [(self.trUtf8(u"Formato:"),
                            [0, self.trUtf8(u"Padrão"),
                             self.trUtf8(u"Romano"),
                             self.trUtf8(u"Extenso")])]
                result = fedit(options,
                               title=self.trUtf8(u"Formato da Data"),
                               icon=QtGui.QIcon(":/ecolog.png"),
                               parent=self)
                if result is not None:
                    formato = result[0] + 1
                    report = Report(filename=filename,
                                    projectdata=self.projectdata,
                                    headers=sheet.headers,
                                    types=sheet.types,
                                    data=sheet.data,
                                    filter=filter_expr)
                    report.Label(formato)
                    self.mdi.addSubWindow(report)
                    self.mdi.setWindowIcon(QtGui.QIcon(":/web.png"))
                    report.show()

    def report_checklist(self):
        dataset = self.projectdata["dataset1"]
        window = self.find_mdi_child(dataset)
        if window:
            self.mdi.setActiveSubWindow(window)
            sheet = self.active_mdi_child()
            filter_expr = self.get_filter(sheet)
            filename = QtGui.QFileDialog.getSaveFileName(self,
                                                         self.trUtf8(u"Salvar Como"), os.getcwd(),
                                                         self.trUtf8(u"Relatórios (*.htm *.html)"))
            if not filename.isEmpty():
                report = Report(filename=filename,
                                projectdata=self.projectdata,
                                headers=sheet.headers,
                                types=sheet.types,
                                data=sheet.data,
                                filter=filter_expr)
                report.Geral()
                self.mdi.addSubWindow(report)
                self.mdi.setWindowIcon(QtGui.QIcon(":/web.png"))
                report.show()

    def report_stats(self):
        dataset = self.projectdata["dataset1"]
        window = self.find_mdi_child(dataset)
        if window:
            self.mdi.setActiveSubWindow(window)
            sheet = self.active_mdi_child()
            filename = QtGui.QFileDialog.getSaveFileName(self,
                                                         self.trUtf8(u"Salvar Como"), os.getcwd(),
                                                         self.trUtf8(u"Relatórios (*.htm *.html)"))
            if not filename.isEmpty():
                if self.projectdata.has_key("dataset2") and len(self.projectdata["dataset2"]) > 0:
                    options = [(self.trUtf8(u"Estatísticas:"),
                                [0, self.trUtf8(u"Famílias"),
                                 self.trUtf8(u"Gêneros"),
                                 self.trUtf8(u"Espécies"),
                                 self.trUtf8("Locais"),
                                 self.trUtf8("Descritores"),
                                 self.trUtf8(u"Variáveis"),
                                 self.trUtf8(u"Sequências")])]
                else:
                    options = [(self.trUtf8(u"Estatísticas:"),
                                [0, self.trUtf8(u"Famílias"),
                                 self.trUtf8(u"Gêneros"),
                                 self.trUtf8(u"Espécies"),
                                 self.trUtf8("Locais"),
                                 self.trUtf8("Descritores"),
                                 self.trUtf8(u"Sequências")])]
                result = fedit(options,
                               title=self.trUtf8(u"Relatório Estatístico"),
                               icon=QtGui.QIcon(":/ecolog.png"),
                               parent=self)
                if result is not None:
                    graph_it = False
                    option = result[0] + 1
                    if option < 5:
                        graph_it = True

                    if option == 5:
                        descriptors = sheet.headers[12:len(sheet.headers)]
                        if len(descriptors) == 0:
                            QtGui.QMessageBox.warning(self, self.trUtf8(u"Aviso"),
                                                      self.trUtf8(u"Não há descritores na planilha"))
                            return

                    if option == 6:
                        if not self.projectdata.has_key("dataset2") or self.projectdata["dataset2"] is None:
                            QtGui.QMessageBox.warning(self, self.trUtf8(u"Aviso"),
                                                      self.trUtf8(u"O projeto não contém uma planilha de variáveis"))
                            return
                        else:
                            dataset = self.projectdata["dataset2"]
                            window = self.find_mdi_child(dataset)
                            if window:
                                self.mdi.setActiveSubWindow(window)
                                sheet = self.active_mdi_child()
                                variables = sheet.headers[4:len(sheet.headers)]
                                if len(variables) == 0:
                                    QtGui.QMessageBox.warning(self, self.trUtf8(u"Aviso"),
                                                              self.trUtf8(u"Não há variáveis na planilha"))
                                    return

                    if option == 7:
                        if not any("SEQ" in s.upper() for s in sheet.headers):
                            QtGui.QMessageBox.warning(self, self.trUtf8(u"Aviso"),
                                                      self.trUtf8(u"Não há sequências na planilha"))
                            return
                    report = Report(filename=filename,
                                    projectdata=self.projectdata,
                                    headers=sheet.headers,
                                    types=sheet.types,
                                    data=sheet.data,
                                    filter="")
                    report.Stats(option, graph_it)
                    self.mdi.addSubWindow(report)
                    self.mdi.setWindowIcon(QtGui.QIcon(":/web.png"))
                    report.show()

    def report_names(self):
        dataset = self.projectdata["dataset1"]
        window = self.find_mdi_child(dataset)
        if window:
            self.mdi.setActiveSubWindow(window)
            sheet = self.active_mdi_child()
            filename = QtGui.QFileDialog.getSaveFileName(self,
                                                         self.trUtf8(u"Salvar Como"), os.getcwd(),
                                                         self.trUtf8(u"Relatórios (*.htm *.html)"))
            if not filename.isEmpty():
                report = Report(filename=filename,
                                projectdata=self.projectdata,
                                headers=sheet.headers,
                                types=sheet.types,
                                data=sheet.data,
                                filter="")
                report.Check()
                self.mdi.addSubWindow(report)
                self.mdi.setWindowIcon(QtGui.QIcon(":/web.png"))
                report.show()

    def report_geocode(self):
        dataset = self.projectdata["dataset1"]
        window = self.find_mdi_child(dataset)
        if window:
            self.mdi.setActiveSubWindow(window)
            sheet = self.active_mdi_child()
            filename = QtGui.QFileDialog.getSaveFileName(self,
                                                         self.trUtf8(u"Salvar Como"), os.getcwd(),
                                                         self.trUtf8(u"Relatórios (*.htm *.html)"))
            if not filename.isEmpty():
                report = Report(filename=filename,
                                projectdata=self.projectdata,
                                headers=sheet.headers,
                                types=sheet.types,
                                data=sheet.data,
                                filter="")
                report.Georef()
                self.mdi.addSubWindow(report)
                self.mdi.setWindowIcon(QtGui.QIcon(":/web.png"))
                report.show()

    def analysis_diversity(self):
        dataset1 = self.projectdata["dataset1"]
        window = self.find_mdi_child(dataset1)
        if window:
            self.mdi.setActiveSubWindow(window)
            sheet = self.active_mdi_child()
            datmat = sheet.gridToArray(field1=0, field2=3, kind=2)
            filename = QtGui.QFileDialog.getSaveFileName(self,
                                                         self.trUtf8(u"Salvar Como"), os.getcwd(),
                                                         self.trUtf8(u"Relatórios (*.htm *.html)"))
            if not filename.isEmpty():
                graph_it = True
                report = Report(filename=filename,
                                projectdata=self.projectdata,
                                headers=sheet.headers,
                                types=sheet.types,
                                data=sheet.data,
                                filter="")
                report.Divers(datmat, graph_it)
                self.mdi.addSubWindow(report)
                self.mdi.setWindowIcon(QtGui.QIcon(":/web.png"))
                report.show()

    def analysis_cluster(self):
        options = [(self.trUtf8(u"Transformação:"),
                    [0, self.trUtf8(u"Nenhuma"),
                     self.trUtf8(u"Logaritmo comum (log 10)"),
                     self.trUtf8(u"Logaritmo natural (log e)"),
                     self.trUtf8(u"Raiz quadrada"),
                     self.trUtf8(u"Arcosseno")]),
                   (self.trUtf8(u"Coeficiente:"),
                    [0, self.trUtf8(u"Bray-Curtis"),
                     self.trUtf8(u"Canberra"),
                     self.trUtf8(u"Manhattan"),
                     self.trUtf8(u"Euclidiana"),
                     self.trUtf8(u"Euclidiana normalizada"),
                     self.trUtf8(u"Euclidiana quadrada"),
                     self.trUtf8(u"Morisita-Horn"),
                     self.trUtf8(u"Correlação"),
                     self.trUtf8(u"Jaccard"),
                     self.trUtf8(u"Dice-Sorenson"),
                     self.trUtf8(u"Kulczynski"),
                     self.trUtf8(u"Ochiai")]),
                   (self.trUtf8(u"Método:"),
                    [2, self.trUtf8(u"Ligação simples (SLM)"),
                     self.trUtf8(u"Ligação completa (CLM)"),
                     self.trUtf8(u"Ligação média (UPGMA)"),
                     self.trUtf8(u"Ligação ponderada (WPGMA)"),
                     self.trUtf8(u"Centroide (UPGMC)"),
                     self.trUtf8(u"Mediana (WPGMC)"),
                     self.trUtf8(u"Método de Ward")])]
        result = fedit(options,
                       title=self.trUtf8(u"Métodos de Agrupamento"),
                       icon=QtGui.QIcon(":/ecolog.png"),
                       parent=self)
        if result is not None:
            transf = result[0]
            coef = result[1]
            tipo = iif(coef > 6, 1, 2)
            if tipo == 1: transf = 0
            method = result[2]
            dataset1 = self.projectdata["dataset1"]
            window = self.find_mdi_child(dataset1)
            if window:
                self.mdi.setActiveSubWindow(window)
                sheet = self.active_mdi_child()
                datmat = sheet.gridToArray(field1=0, field2=3, kind=tipo)
                filename = QtGui.QFileDialog.getSaveFileName(self,
                                                             self.trUtf8(u"Salvar Como"), os.getcwd(),
                                                             self.trUtf8(u"Relatórios (*.htm *.html)"))
                if not filename.isEmpty():
                    graph_it = True
                    report = Report(filename=filename,
                                    projectdata=self.projectdata,
                                    headers=sheet.headers,
                                    types=sheet.types,
                                    data=sheet.data,
                                    filter="")
                    report.Cluster(datmat, transf, coef, method, graph_it)
                    self.mdi.addSubWindow(report)
                    self.mdi.setWindowIcon(QtGui.QIcon(":/web.png"))
                    report.show()

    def ordination_pca(self):
        options = [(self.trUtf8(u"Centragem"), False),
                   (self.trUtf8(u"Estandardização"), False),
                   (self.trUtf8(u"Transformação:"),
                    [0, self.trUtf8(u"Nenhuma"),
                     self.trUtf8(u"Logaritmo comum (log 10)"),
                     self.trUtf8(u"Logaritmo natural (log e)"),
                     self.trUtf8(u"Raiz quadrada"),
                     self.trUtf8(u"Arcosseno")]),
                   (self.trUtf8(u"Matriz:"),
                    [1, self.trUtf8(u"Covariância"),
                     self.trUtf8(u"Correlação")])]
        result = fedit(options,
                       title=self.trUtf8(u"Opções"),
                       icon=QtGui.QIcon(":/ecolog.png"),
                       parent=self)
        if result is not None:
            center = result[0]
            scale = result[1]
            transf = result[2]
            index = result[3] + 1
            dataset1 = self.projectdata["dataset1"]
            window = self.find_mdi_child(dataset1)
            if window:
                self.mdi.setActiveSubWindow(window)
                sheet = self.active_mdi_child()
                datmat = sheet.gridToArray(field1=0, field2=3, kind=2)
                filename = QtGui.QFileDialog.getSaveFileName(self,
                                                             self.trUtf8(u"Salvar Como"), os.getcwd(),
                                                             self.trUtf8(u"Relatórios (*.htm *.html)"))
                if not filename.isEmpty():
                    method = 0
                    graph_it = True
                    report = Report(filename=filename,
                                    projectdata=self.projectdata,
                                    headers=sheet.headers,
                                    types=sheet.types,
                                    data=sheet.data,
                                    filter="")
                    report.Ord(datmat, None, transf, method, -1, index, center, scale, 0, "", 0, [], graph_it)
                    self.mdi.addSubWindow(report)
                    self.mdi.setWindowIcon(QtGui.QIcon(":/web.png"))
                    report.show()

    def ordination_pcoa(self):
        options = [(self.trUtf8(u"Transformação:"),
                    [0, self.trUtf8(u"Nenhuma"),
                     self.trUtf8(u"Logaritmo comum (log 10)"),
                     self.trUtf8(u"Logaritmo natural (log e)"),
                     self.trUtf8(u"Raiz quadrada"),
                     self.trUtf8(u"Arcosseno")]),
                   (self.trUtf8(u"Coeficiente:"),
                    [3, self.trUtf8(u"Bray-Curtis"),
                     self.trUtf8(u"Canberra"),
                     self.trUtf8(u"Manhattan"),
                     self.trUtf8(u"Euclidiana"),
                     self.trUtf8(u"Euclidiana normalizada"),
                     self.trUtf8(u"Euclidiana quadrada"),
                     self.trUtf8(u"Morisita-Horn")])]
        result = fedit(options,
                       title=self.trUtf8(u"Opções"),
                       icon=QtGui.QIcon(":/ecolog.png"),
                       parent=self)
        if result is not None:
            transf = result[0]
            coef = result[1]
            dataset1 = self.projectdata["dataset1"]
            window = self.find_mdi_child(dataset1)
            if window:
                self.mdi.setActiveSubWindow(window)
                sheet = self.active_mdi_child()
                datmat = sheet.gridToArray(field1=0, field2=3, kind=2)
                filename = QtGui.QFileDialog.getSaveFileName(self,
                                                             self.trUtf8(u"Salvar Como"), os.getcwd(),
                                                             self.trUtf8(u"Relatórios (*.htm *.html)"))
                if not filename.isEmpty():
                    method = 1
                    graph_it = True
                    report = Report(filename=filename,
                                    projectdata=self.projectdata,
                                    headers=sheet.headers,
                                    types=sheet.types,
                                    data=sheet.data,
                                    filter="")
                    report.Ord(datmat, None, transf, method, coef, -1, False, False, 0, "", 0, [], graph_it)
                    self.mdi.addSubWindow(report)
                    self.mdi.setWindowIcon(QtGui.QIcon(":/web.png"))
                    report.show()

    def ordination_nmds(self):
        options = [(self.trUtf8(u"Transformação:"),
                    [0, self.trUtf8(u"Nenhuma"),
                     self.trUtf8(u"Logaritmo comum (log 10)"),
                     self.trUtf8(u"Logaritmo natural (log e)"),
                     self.trUtf8(u"Raiz quadrada"),
                     self.trUtf8(u"Arcosseno")]),
                   (self.trUtf8(u"Coeficiente:"),
                    [3, self.trUtf8(u"Bray-Curtis"),
                     self.trUtf8(u"Canberra"),
                     self.trUtf8(u"Manhattan"),
                     self.trUtf8(u"Euclidiana"),
                     self.trUtf8(u"Euclidiana normalizada"),
                     self.trUtf8(u"Euclidiana quadrada"),
                     self.trUtf8(u"Morisita-Horn")]),
                   (self.trUtf8(u"Número de Interações:"), 50),
                   (self.trUtf8(u"Configuração Inicial:"),
                    [0, self.trUtf8(u"Coordenadas Principais"),
                     self.trUtf8(u"Aleatória")])]
        result = fedit(options,
                       title=self.trUtf8(u"Opções"),
                       icon=QtGui.QIcon(":/ecolog.png"),
                       parent=self)
        if result is not None:
            transf = result[0]
            coef = result[1]
            iter = result[2]
            config = iif(result[3] == 0, "pcoa", "random")
            dataset1 = self.projectdata["dataset1"]
            window = self.find_mdi_child(dataset1)
            if window:
                self.mdi.setActiveSubWindow(window)
                sheet = self.active_mdi_child()
                datmat = sheet.gridToArray(field1=0, field2=3, kind=2)
                filename = QtGui.QFileDialog.getSaveFileName(self,
                                                             self.trUtf8(u"Salvar Como"), os.getcwd(),
                                                             self.trUtf8(u"Relatórios (*.htm *.html)"))
                if not filename.isEmpty():
                    method = 2
                    graph_it = True
                    report = Report(filename=filename,
                                    projectdata=self.projectdata,
                                    headers=sheet.headers,
                                    types=sheet.types,
                                    data=sheet.data,
                                    filter="")
                    report.Ord(datmat, None, transf, method, coef, -1, False, False, iter, config, 0, [], graph_it)
                    self.mdi.addSubWindow(report)
                    self.mdi.setWindowIcon(QtGui.QIcon(":/web.png"))
                    report.show()

    def ordination_ca(self):
        options = [(self.trUtf8(u"Transformação:"),
                    [0, self.trUtf8(u"Nenhuma"),
                     self.trUtf8(u"Logaritmo comum (log 10)"),
                     self.trUtf8(u"Logaritmo natural (log e)"),
                     self.trUtf8(u"Raiz quadrada"),
                     self.trUtf8(u"Arcosseno")]),
                   (self.trUtf8(u"Ordenação:"),
                    [0, self.trUtf8(u"Amostras"),
                     self.trUtf8(u"Espécies")])]
        result = fedit(options,
                       title=self.trUtf8(u"Opções"),
                       icon=QtGui.QIcon(":/ecolog.png"),
                       parent=self)
        if result is not None:
            transf = result[0]
            scaling = result[1] + 1
            dataset1 = self.projectdata["dataset1"]
            window = self.find_mdi_child(dataset1)
            if window:
                self.mdi.setActiveSubWindow(window)
                sheet = self.active_mdi_child()
                datmat = sheet.gridToArray(field1=0, field2=3, kind=2)
                filename = QtGui.QFileDialog.getSaveFileName(self,
                                                             self.trUtf8(u"Salvar Como"), os.getcwd(),
                                                             self.trUtf8(u"Relatórios (*.htm *.html)"))
                if not filename.isEmpty():
                    method = 3
                    graph_it = True
                    report = Report(filename=filename,
                                    projectdata=self.projectdata,
                                    headers=sheet.headers,
                                    types=sheet.types,
                                    data=sheet.data,
                                    filter="")
                    report.Ord(datmat, None, transf, method, -1, -1, False, False, 0, "", scaling, [], graph_it)
                    self.mdi.addSubWindow(report)
                    self.mdi.setWindowIcon(QtGui.QIcon(":/web.png"))
                    report.show()

    def ordination_rda(self):
        if self.projectdata.has_key("dataset2") and len(self.projectdata["dataset2"]) > 0:
            options = [(self.trUtf8(u"Transformação:"),
                        [0, self.trUtf8(u"Nenhuma"),
                         self.trUtf8(u"Logaritmo comum (log 10)"),
                         self.trUtf8(u"Logaritmo natural (log e)"),
                         self.trUtf8(u"Raiz quadrada"),
                         self.trUtf8(u"Arcosseno")])]
            result = fedit(options,
                           title=self.trUtf8(u"Opções"),
                           icon=QtGui.QIcon(":/ecolog.png"),
                           parent=self)
            if result is not None:
                transf = result[0]
                dataset1 = self.projectdata["dataset1"]
                dataset2 = self.projectdata["dataset2"]
                window1 = self.find_mdi_child(dataset1)
                if window1:
                    self.mdi.setActiveSubWindow(window1)
                    sheet1 = self.active_mdi_child()
                    datmat = sheet1.gridToArray(field1=0, field2=3, kind=2)
                    n1, m1 = datmat.shape
                    window2 = self.find_mdi_child(dataset2)
                    if window2:
                        self.mdi.setActiveSubWindow(window2)
                        sheet2 = self.active_mdi_child()
                        envmat = sheet2.gridToArray(field1=0, field2=None, kind=3)
                        n2, m2 = envmat.shape
                        if n1 != n2:
                            QtGui.QMessageBox.warning(self, self.trUtf8(u"Aviso"),
                                                      self.trUtf8(u"As matrizes devem ter o mesmo número de linhas (amostras)"))
                            return
                        if n1 < m2:
                            QtGui.QMessageBox.warning(self, self.trUtf8(u"Aviso"),
                                                      self.trUtf8(u"A matriz de dados não pode ter menos linhas do que as colunas da matriz de variáveis ambientais"))
                            return
                        choice = ChoiceDialog(self.trUtf8(u"Variáveis"),
                                              stringlist=sheet2.headers[1:],
                                              multi=True,
                                              icon=QtGui.QIcon(":/ecolog.png"),
                                              parent=self)
                        if choice.exec_():
                            mask = []
                            for index in range(choice.listWidget.count()):
                                item = choice.listWidget.item(index)
                                if choice.listWidget.isItemSelected(item):
                                    mask.append(index)
                            filename = QtGui.QFileDialog.getSaveFileName(self,
                                                                         self.trUtf8(u"Salvar Como"), os.getcwd(),
                                                                         self.trUtf8(u"Relatórios (*.htm *.html)"))
                            if not filename.isEmpty():
                                method = 4
                                graph_it = True
                                report = Report(filename=filename,
                                                projectdata=self.projectdata,
                                                headers=sheet1.headers,
                                                types=sheet1.types,
                                                data=sheet1.data,
                                                filter="")
                                report.Ord(datmat, envmat, transf, method, -1, -1, False, False, 0, "", 0, mask, graph_it)
                                self.mdi.addSubWindow(report)
                                self.mdi.setWindowIcon(QtGui.QIcon(":/web.png"))
                                report.show()
        else:
            QtGui.QMessageBox.warning(self, self.trUtf8(u"Aviso"),
                                      self.trUtf8(u"O projeto não inclui um arquivo de vsriáveis ambientais"))

    def ordination_cca(self):
        if self.projectdata.has_key("dataset2") and len(self.projectdata["dataset2"]) > 0:
            options = [(self.trUtf8(u"Transformação:"),
                        [0, self.trUtf8(u"Nenhuma"),
                         self.trUtf8(u"Logaritmo comum (log 10)"),
                         self.trUtf8(u"Logaritmo natural (log e)"),
                         self.trUtf8(u"Raiz quadrada"),
                         self.trUtf8(u"Arcosseno")]),
                       (self.trUtf8(u"Ordenação:"),
                        [0, self.trUtf8(u"Amostras"),
                         self.trUtf8(u"Espécies")])]
            result = fedit(options,
                           title=self.trUtf8(u"Opções"),
                           icon=QtGui.QIcon(":/ecolog.png"),
                           parent=self)
            if result is not None:
                transf = result[0]
                scaling = result[1] + 1
                dataset1 = self.projectdata["dataset1"]
                dataset2 = self.projectdata["dataset2"]
                window1 = self.find_mdi_child(dataset1)
                if window1:
                    self.mdi.setActiveSubWindow(window1)
                    sheet1 = self.active_mdi_child()
                    datmat = sheet1.gridToArray(field1=0, field2=3, kind=2)
                    n1, m1 = datmat.shape
                    window2 = self.find_mdi_child(dataset2)
                    if window2:
                        self.mdi.setActiveSubWindow(window2)
                        sheet2 = self.active_mdi_child()
                        envmat = sheet2.gridToArray(field1=0, field2=None, kind=3)
                        n2, m2 = envmat.shape
                        if n1 != n2:
                            QtGui.QMessageBox.warning(self, self.trUtf8(u"Aviso"),
                                                      self.trUtf8(u"As matrizes devem ter o mesmo número de linhas (amostras)"))
                            return
                        choice = ChoiceDialog(self.trUtf8(u"Variáveis"),
                                              stringlist=sheet2.headers[1:],
                                              multi=True,
                                              icon=QtGui.QIcon(":/ecolog.png"),
                                              parent=self)
                        if choice.exec_():
                            mask = []
                            for index in range(choice.listWidget.count()):
                                item = choice.listWidget.item(index)
                                if choice.listWidget.isItemSelected(item):
                                    mask.append(index)
                            filename = QtGui.QFileDialog.getSaveFileName(self,
                                                                         self.trUtf8(u"Salvar Como"), os.getcwd(),
                                                                         self.trUtf8(u"Relatórios (*.htm *.html)"))
                            if not filename.isEmpty():
                                method = 5
                                graph_it = True
                                report = Report(filename=filename,
                                                projectdata=self.projectdata,
                                                headers=sheet1.headers,
                                                types=sheet1.types,
                                                data=sheet1.data,
                                                filter="")
                                report.Ord(datmat, envmat, transf, method, -1, -1, False, False, 0, "", scaling, mask, graph_it)
                                self.mdi.addSubWindow(report)
                                self.mdi.setWindowIcon(QtGui.QIcon(":/web.png"))
                                report.show()
        else:
            QtGui.QMessageBox.warning(self, self.trUtf8(u"Aviso"),
                                      self.trUtf8(u"O projeto não inclui um arquivo de vsriáveis ambientais"))

    def help_about(self):
        QtGui.QMessageBox.about(self, self.trUtf8(u"Sobre ECOLOG"),
                                "<b>ECOLOG</b> v " + __version__ + '\n' + \
                                "<p>" + self.trUtf8("Sistema de Banco de Dados para Levantamentos Ecológicos de Campo") + '\n' + \
                                "<p>Copyright &copy; 1990-2016 Mauro J. Cavalcanti" + '\n' + \
                                "<p>Python: " + platform.python_version() + '\n' + \
                                "<br>Qt: " + QT_VERSION_STR + '\n' + \
                                "<br>PyQt: " +  PYQT_VERSION_STR + '\n' + \
                                "<br>SQLite: " + sqlite3.sqlite_version + '\n' + \
                                "<br>PySQLite: " + sqlite3.version + '\n' + \
                                "<br>" + self.trUtf8("Plataforma: ") + platform.system() + ' ' + platform.release() + '\n' + \
                                "<p><a href=""http://ecolog.sourceforge.net"">" + self.trUtf8("Página do ECOLOG na Internet") + "</a>"
                                )

    def help_about_qt(self):
        QtGui.QMessageBox.aboutQt(self)

    def update_menus(self):
        enable = (self.active_mdi_child() is not None)
        self.fileEditAction.setEnabled(enable)
        self.fileReloadAction.setEnabled(enable)
        self.fileCloseAction.setEnabled(enable)
        self.fileAddAction.setEnabled(enable)
        self.fileRemoveAction.setEnabled(enable)
        self.fileImportAction.setEnabled(enable)
        self.fileExportAction.setEnabled(enable)
        self.dataSortAction.setEnabled(enable)
        self.dataFindAction.setEnabled(enable)
        self.dataFilterAction.setEnabled(enable)
        self.reportCatalogAction.setEnabled(enable)
        self.reportLabelsAction.setEnabled(enable)
        self.reportChecklistAction.setEnabled(enable)
        self.reportStatsAction.setEnabled(enable)
        self.reportNamesAction.setEnabled(enable)
        self.reportGeocodeAction.setEnabled(enable)
        self.analysisDiversityAction.setEnabled(enable)
        self.analysisClusterAction.setEnabled(enable)
        self.ordinationMenu.setEnabled(enable)
        self.ordinationPCAAction.setEnabled(enable)
        self.ordinationPCOAction.setEnabled(enable)
        self.ordinationNMDSAction.setEnabled(enable)
        self.ordinationCAAction.setEnabled(enable)
        self.ordinationRDAAction.setEnabled(enable)
        self.ordinationCCAction.setEnabled(enable)
        self.windowNextAction.setEnabled(enable)
        self.windowPrevAction.setEnabled(enable)
        self.windowCascadeAction.setEnabled(enable)
        self.windowTileVerticalAction.setEnabled(enable)
        self.windowTileHorizontalAction.setEnabled(enable)
        self.windowCloseAction.setEnabled(enable)
        self.toolbtn.setEnabled(enable)
        if enable:
            self.setWindowTitle("ECOLOG - " + basename(self.projectfile))
        else:
            self.setWindowTitle("ECOLOG")

    def update_window_menu(self):
        self.windowMenu.clear()
        self.add_actions(self.windowMenu, (self.windowNextAction,
                                           self.windowPrevAction, None, self.windowCascadeAction,
                                           self.windowTileVerticalAction, self.windowTileHorizontalAction,
                                           None,
                                           self.windowCloseAction))
        windows = self.mdi.subWindowList()
        if not windows:
            return
        self.windowMenu.addSeparator()
        menu = self.windowMenu
        for i, window in enumerate(windows):
            child = window.widget()
            title = "%d %s" % (i + 1, window.windowTitle())
            if i < 9:
                title = '&' + title
            action = self.windowMenu.addAction(title)
            action.setCheckable(True)
            action.setChecked(child is self.active_mdi_child())
            action.triggered.connect(self.windowMapper.map)
            self.windowMapper.setMapping(action, window)

    def create_actions(self):
        self.fileNewAction = self.create_action(self.trUtf8(u"&Novo..."), self.file_new,
                                                QtGui.QKeySequence.New, "new", self.trUtf8(u"Criar um novo projeto"))
        self.fileOpenAction = self.create_action(self.trUtf8(u"&Abrir..."), self.file_open,
                                                 QtGui.QKeySequence.Open, "open", self.trUtf8("Abrir um projeto existente"))
        self.fileEditAction = self.create_action(self.trUtf8(u"&Modificar..."), self.file_edit,
                                                 None, "edit", self.trUtf8(u"Modificar o projeto atual"))
        self.fileReloadAction = self.create_action(self.trUtf8(u"&Recarregar"), self.file_reload,
                                                   None, "reload", self.trUtf8(u"Recarregar o projeto atual"))
        self.fileCloseAction = self.create_action(self.trUtf8(u"Fechar"), self.file_close,
                                                  None, "close", self.trUtf8(u"Fechar o projeto atual"))
        self.fileAddAction = self.create_action(self.trUtf8(u"A&dicionar..."), self.file_add,
                                                None, "add", self.trUtf8(u"Adicionar arquivo ao projeto atual"))
        self.fileRemoveAction = self.create_action(self.trUtf8(u"&Remover..."), self.file_remove,
                                                   None, "remove", self.trUtf8(u"Remover arquivo do projeto atual"))
        self.fileImportAction = self.create_action(self.trUtf8(u"Importar..."), self.file_import,
                                                   None, "import", self.trUtf8(u"Importar dados externos"))
        self.fileExportAction = self.create_action(self.trUtf8(u"Exportar..."), self.file_export,
                                                   None, "export", self.trUtf8(u"Exportar dados para outros programas"))
        self.fileExitAction = self.create_action(self.trUtf8(u"&Sair"), self.close,
                                                 "Ctrl+Q", "exit", self.trUtf8(u"Encerrar o programa"))

        self.dataSortAction = self.create_action(self.trUtf8(u"&Ordenar"), self.data_sort,
                                                 "Ctrl+S", "sort", self.trUtf8(u"Ordenar registros da tabela atual"))
        self.dataFindAction = self.create_action(self.trUtf8(u"&Pesquisar..."), self.data_find,
                                                 QtGui.QKeySequence.Find, "search", self.trUtf8(u"Pesquisar registros na tabela atual"))
        self.dataFilterAction = self.create_action(self.trUtf8(u"&Filtrar..."), self.data_filter,
                                                   None, "query", self.trUtf8(u"Definir filtro para exibição de registros"))

        self.reportCatalogAction = self.create_action(self.trUtf8(u"&Catálogo..."), self.report_catalog,
                                                      None, "catalog", self.trUtf8(u"Catálogo das coletas"))
        self.reportLabelsAction = self.create_action(self.trUtf8(u"&Etiquetas..."), self.report_labels,
                                                     None, "labels", self.trUtf8(u"Etiquetas de coleta"))
        self.reportChecklistAction = self.create_action(self.trUtf8(u"&Geral..."), self.report_checklist,
                                                        None, "geral", self.trUtf8(u"Relatório geral de espécies"))
        self.reportStatsAction = self.create_action(self.trUtf8(u"E&statísticas..."), self.report_stats,
                                                    None, "stat", self.trUtf8(u"Relatório estatístico"))
        self.reportNamesAction = self.create_action(self.trUtf8(u"&Nomenclatura..."), self.report_names,
                                                    None, "check", self.trUtf8(u"Verificar nomes no Catálogo da Vida (www.catalogueoflife.org)"))
        self.reportGeocodeAction = self.create_action(self.trUtf8(u"&Geocodificação..."), self.report_geocode,
                                                      None, "map", self.trUtf8(u"Verificar coordenadas geográficas no Google Maps (maps.google.com)"))

        self.analysisDiversityAction = self.create_action(self.trUtf8(u"Diversidade"), self.analysis_diversity,
                                                          None, "divers", self.trUtf8(u"Análise de diversidade"))
        self.analysisClusterAction = self.create_action(self.trUtf8(u"Agrupamento..."), self.analysis_cluster,
                                                        None, "clus", self.trUtf8(u"Análise de agrupamentos"))
        self.ordinationPCAAction = self.create_action(self.trUtf8(u"Análise de Componentes Principais..."),
                                                      self.ordination_pca, None, None, self.trUtf8(u"Análise de componentes principais (PCA)"))
        self.ordinationPCOAction = self.create_action(self.trUtf8(u"Análise de Coordenadas Principais..."),
                                                      self.ordination_pcoa, None, None, self.trUtf8(u"Análise de coordenadas principais (PCOA)"))
        self.ordinationNMDSAction = self.create_action(self.trUtf8(u"Escalonamento Multidimensional Não-Métrico..."),
                                                       self.ordination_nmds, None, None, self.trUtf8(u"Escalonamento multidimensional não-métrico (NMDS)"))
        self.ordinationCAAction = self.create_action(self.trUtf8(u"Análise de Correspondências..."),
                                                     self.ordination_ca, None, None, self.trUtf8(u"Análise de correspondências (CA)"))
        self.ordinationRDAAction = self.create_action(self.trUtf8(u"Análise de Redundâncias..."),
                                                      self.ordination_rda, None, None, self.trUtf8(u"Análise de redundâncias (RDA)"))
        self.ordinationCCAction = self.create_action(self.trUtf8(u"Análise de Correspondências Canônica..."),
                                                     self.ordination_cca, None, None, self.trUtf8(u"Análise de correspondências canônica (CCA)"))

        self.windowNextAction = self.create_action(self.trUtf8(u"&Próxima"),
                                                   self.mdi.activateNextSubWindow, QtGui.QKeySequence.NextChild, "forward")
        self.windowPrevAction = self.create_action(self.trUtf8(u"&Anterior"),
                                                   self.mdi.activatePreviousSubWindow, QtGui.QKeySequence.PreviousChild, "back")
        self.windowCascadeAction = self.create_action(self.trUtf8(u"Em &Cascata"),
                                                      self.mdi.cascadeSubWindows, None, "cascade")
        self.windowTileVerticalAction = self.create_action(self.trUtf8(u"Lado a Lado &Vertical"),
                                                           self.mdi.tileSubWindows, None, "tilever")
        self.windowTileHorizontalAction = self.create_action(self.trUtf8(u"Lado a Lado &Horizontal"),
                                                             self.window_tile_horizontal, None, "tilehor")
        self.windowCloseAction = self.create_action(self.trUtf8(u"&Fechar"),
                                                    self.mdi.closeActiveSubWindow, QtGui.QKeySequence.Close, "stop")

        self.helpAboutAction = self.create_action(self.trUtf8(u"Sobre ECOLOG"), self.help_about,
                                                  None, "about", self.trUtf8(u"Exibir informação sobre o programa"))
        self.helpAboutQtAction = self.create_action(self.trUtf8(u"Sobre Qt"), self.help_about_qt,
                                                    None, "qt", self.trUtf8(u"Exibir informação sobre a interface"))

    def create_menus(self):
        fileMenu = self.menuBar().addMenu(self.trUtf8(u"&Projeto"))
        self.add_actions(fileMenu, (self.fileNewAction, self.fileOpenAction, self.fileEditAction,
                                    self.fileReloadAction, self.fileCloseAction, None,
                                    self.fileAddAction, self.fileRemoveAction,
                                    None, self.fileImportAction, self.fileExportAction,
                                    None, self.fileExitAction))

        dataMenu = self.menuBar().addMenu(self.trUtf8(u"&Dados"))
        self.add_actions(dataMenu, (self.dataSortAction, self.dataFindAction, None, self.dataFilterAction))

        reportMenu = self.menuBar().addMenu(self.trUtf8(u"Relatórios"))
        self.add_actions(reportMenu, (self.reportCatalogAction, self.reportLabelsAction,
                                      self.reportChecklistAction, self.reportStatsAction,
                                      None, self.reportNamesAction, self.reportGeocodeAction))

        analysisMenu = self.menuBar().addMenu(self.trUtf8(u"&Análises"))
        self.add_actions(analysisMenu, (self.analysisDiversityAction, self.analysisClusterAction))
        self.ordinationMenu = analysisMenu.addMenu(QtGui.QIcon(":/ord.png"), self.trUtf8(u"Ordenação"))
        self.add_actions(self.ordinationMenu, (self.ordinationPCAAction, self.ordinationPCOAction,
                                               self.ordinationNMDSAction, self.ordinationCAAction,
                                               self.ordinationRDAAction, self.ordinationCCAction))

        self.windowMenu = self.menuBar().addMenu(self.trUtf8(u"&Janela"))
        self.connect(self.windowMenu, QtCore.SIGNAL("aboutToShow()"), self.update_window_menu)

        helpMenu = self.menuBar().addMenu(self.trUtf8(u"A&juda"))
        self.add_actions(helpMenu, (self.helpAboutAction, self.helpAboutQtAction))

    def create_toolbars(self):
        toolbar = self.addToolBar("Toolbar")
        toolbar.setObjectName("Toolbar")
        self.toolbtn = QtGui.QToolButton(toolbar)
        self.toolbtn.setMenu(self.ordinationMenu)
        self.toolbtn.setIcon(QtGui.QIcon(":/ord.png"))
        self.toolbtn.setPopupMode(QtGui.QToolButton.MenuButtonPopup)
        self.add_actions(toolbar, (self.fileNewAction, self.fileOpenAction, self.fileReloadAction,
                                   self.fileAddAction, self.fileRemoveAction, None,
                                   self.dataSortAction, self.dataFindAction, self.dataFilterAction, None,
                                   self.reportCatalogAction, self.reportLabelsAction, self.reportChecklistAction,
                                   self.reportStatsAction, self.reportNamesAction, self.reportGeocodeAction,
                                   None, self.analysisDiversityAction, self.analysisClusterAction))
        toolbar.addWidget(self.toolbtn)

    def create_statusbar(self):
        self.statusBar().showMessage(self.trUtf8(u"Pronto"))

    def create_action(self, text, slot=None, shortcut=None, icon=None,
                      tip=None, checkable=False, signal="triggered()"):
        action = QtGui.QAction(text, self)
        if icon is not None:
            action.setIcon(QtGui.QIcon(":/{0}.png".format(icon)))
        if shortcut is not None:
            action.setShortcut(shortcut)
        if tip is not None:
            action.setToolTip(tip)
            action.setStatusTip(tip)
        if slot is not None:
            self.connect(action, QtCore.SIGNAL(signal), slot)
        if checkable:
            action.setCheckable(True)
        return action

    def add_actions(self, target, actions):
        for action in actions:
            if action is None:
                target.addSeparator()
            else:
                target.addAction(action)

    def restore_settings(self):
        settings = QtCore.QSettings()
        size = settings.value("MainWindow/Size",
                              QtCore.QVariant(QtCore.QSize(766, 485))).toSize()
        self.resize(size)
        position = settings.value("MainWindow/Position",
                                  QtCore.QVariant(QtCore.QPoint(192, 107))).toPoint()
        self.move(position)
        self.restoreState(
            settings.value("MainWindow/State").toByteArray())

    def save_settings(self):
        settings = QtCore.QSettings()
        settings.setValue("MainWindow/Size", QtCore.QVariant(self.size()))
        settings.setValue("MainWindow/Position",
                          QtCore.QVariant(self.pos()))
        settings.setValue("MainWindow/State",
                          QtCore.QVariant(self.saveState()))

    def load_file(self, filename):
        if filename is None: return

        user = None
        pwd = None

        if filename.endswith(".docs"):
            login = LoginDialog(self.trUtf8(u"Autenticação do Usuário"),
                                icon=QtGui.QIcon(":/ecolog.png"),
                                parent=self)
            if login.exec_():
                user = login.username.text()
                pwd = login.password.text()
            else:
                return

        window = Sheet(filename)
        QtGui.QApplication.setOverrideCursor(QtGui.QCursor(QtCore.Qt.WaitCursor))
        err = window.loadData(user, pwd)
        if err is not None:
            QtGui.QApplication.restoreOverrideCursor()
            QtGui.QMessageBox.critical(self, self.trUtf8("Erro"), str(err))
            window.close()
            del window
        else:
            QtGui.QApplication.restoreOverrideCursor()
            window.setData()
            self.mdi.addSubWindow(window)
            self.mdi.setWindowIcon(QtGui.QIcon(":/table.png"))
            window.show()

    def window_tile_horizontal(self):
        if not len(self.mdi.subWindowList()):
            return
        windowheight = float(self.mdi.height()) / len(self.mdi.subWindowList())
        i = 0
        for window in self.mdi.subWindowList():
            window.showNormal()
            window.setGeometry(0, int(windowheight * i),
                               self.mdi.width(), int(windowheight))
            window.raise_()
            i += 1

    def active_mdi_child(self):
        activeSubWindow = self.mdi.activeSubWindow()
        if activeSubWindow:
            return activeSubWindow.widget()
        return None

    def find_mdi_child(self, fileName):
        for window in self.mdi.subWindowList():
            if window.widget().filename == fileName:
                return window
        return None

    def init_project_data(self):
        self.projectdata["title"] = ""
        self.projectdata["author"] = ""
        self.projectdata["description"] = ""
        self.projectdata["method"] = ""
        self.projectdata["size"] = ""
        self.projectdata["country"] = ""
        self.projectdata["state"] = ""
        self.projectdata["province"] = ""
        self.projectdata["locality"] = ""
        self.projectdata["latitude"] = ""
        self.projectdata["longitude"] = ""
        self.projectdata["elevation"] = ""
        self.projectdata["role"] = ""
        self.projectdata["institution"] = ""
        self.projectdata["address1"] = ""
        self.projectdata["address2"] = ""
        self.projectdata["city"] = ""
        self.projectdata["uf"] = ""
        self.projectdata["zip"] = ""
        self.projectdata["phone"] = ""
        self.projectdata["fax"] = ""
        self.projectdata["email"] = ""
        self.projectdata["website"] = ""
        self.projectdata["funding"] = ""

        self.description = [(self.trUtf8(u"Título:"), ""),
                            (self.trUtf8(u"Responsável:"), ""),
                            (self.trUtf8(u"Descrição:"), "\n"),
                            (self.trUtf8(u"Método:"), [self.trUtf8(u"Coleta Aleatória"),
                                                       self.trUtf8(u"Coleta Aleatória"), self.trUtf8(u"Estação"), self.trUtf8(u"Parcela"),
                                                       self.trUtf8(u"Ponto"), self.trUtf8(u"Quadrante"), self.trUtf8(u"Transecção")]),
                            (self.trUtf8(u"Tamanho (unidade):"), "0")]

        self.site = [(self.trUtf8(u"País:"), ""),
                     (self.trUtf8(u"Estado:"), ""),
                     (self.trUtf8(u"Município:"), ""),
                     (self.trUtf8(u"Localidade:"), "\n"),
                     (self.trUtf8(u"Latitude:"), ""),
                     (self.trUtf8(u"Longitude:"), ""),
                     (self.trUtf8(u"Altitude (unidade):"), "")]

        self.address = [(self.trUtf8(u"Função:"), ""),
                        (self.trUtf8(u"Instituição:"), ""),
                        (self.trUtf8(u"Setor:"), ""),
                        (self.trUtf8(u"Endereço:"), ""),
                        (self.trUtf8(u"Cidade:"), ""),
                        (self.trUtf8(u"Estado:"), ""),
                        (self.trUtf8(u"CEP:"), ""),
                        (self.trUtf8(u"Telefone:"), ""),
                        (self.trUtf8(u"Fax:"), ""),
                        (self.trUtf8(u"E-mail:"), ""),
                        (self.trUtf8(u"Website:"), ""),
                        (self.trUtf8(u"Apoio:"), "")]

        self.datagroup = ((self.description, self.trUtf8(u"Descrição"), ""),
                          (self.site, self.trUtf8(u"Localização"), ""),
                          (self.address, self.trUtf8(u"Contato"), ""))

    def get_project_data(self, result):
        self.projectdata["title"] = result[0][0]
        self.projectdata["author"] = result[0][1]
        self.projectdata["description"] = result[0][2]
        self.projectdata["method"] = result[0][3]
        self.projectdata["size"] = result[0][4]
        self.projectdata["country"] = result[1][0]
        self.projectdata["state"] = result[1][1]
        self.projectdata["province"] = result[1][2]
        self.projectdata["locality"] = result[1][3]
        self.projectdata["latitude"] = result[1][4]
        self.projectdata["longitude"] = result[1][5]
        self.projectdata["elevation"] = result[1][6]
        self.projectdata["role"] = result[2][0]
        self.projectdata["institution"] = result[2][1]
        self.projectdata["address1"] = result[2][2]
        self.projectdata["address2"] = result[2][3]
        self.projectdata["city"] = result[2][4]
        self.projectdata["uf"] = result[2][5]
        self.projectdata["zip"] = result[2][6]
        self.projectdata["phone"] = result[2][7]
        self.projectdata["fax"] = result[2][8]
        self.projectdata["email"] = result[2][9]
        self.projectdata["website"] = result[2][10]
        self.projectdata["funding"] = result[2][11]
        self.projectdata["datetime"] = datetime.now()

    def set_project_data(self):
        self.description = [(self.trUtf8(u"Título:"), self.projectdata["title"]),
                            (self.trUtf8(u"Responsável:"), self.projectdata["author"]),
                            (self.trUtf8(u"Descrição:"), self.projectdata["description"] + '\n'),
                            (self.trUtf8(u"Método:"), [self.projectdata["method"],
                                                       self.trUtf8(u"Coleta Aleatória"), self.trUtf8("Estação"),
                                                       self.trUtf8(u"Parcela"), self.trUtf8(u"Ponto"), self.trUtf8(u"Quadrante"),
                                                       self.trUtf8(u"Transecção")]),
                            (self.trUtf8(u"Tamanho (unidade):"), self.projectdata["size"])]

        self.site = [(self.trUtf8(u"País:"), self.projectdata["country"]),
                     (self.trUtf8(u"Estado:"), self.projectdata["state"]),
                     (self.trUtf8(u"Município:"), self.projectdata["province"]),
                     (self.trUtf8(u"Localidade:"), self.projectdata["locality"] + '\n'),
                     (self.trUtf8(u"Latitude:"), self.projectdata["latitude"]),
                     (self.trUtf8(u"Longitude:"), self.projectdata["longitude"]),
                     (self.trUtf8(u"Altitude (unidade):"), self.projectdata["elevation"])]

        self.address = [(self.trUtf8(u"Função:"), self.projectdata["role"]),
                        (self.trUtf8(u"Instituição:"), self.projectdata["institution"]),
                        (self.trUtf8(u"Setor:"), self.projectdata["address1"]),
                        (self.trUtf8(u"Endereço:"), self.projectdata["address2"]),
                        (self.trUtf8(u"Cidade:"), self.projectdata["city"]),
                        (self.trUtf8(u"Estado:"), self.projectdata["uf"]),
                        (self.trUtf8(u"CEP:"), self.projectdata["zip"]),
                        (self.trUtf8(u"Telefone:"), self.projectdata["phone"]),
                        (self.trUtf8(u"Fax:"), self.projectdata["fax"]),
                        (self.trUtf8(u"E-mail:"), self.projectdata["email"]),
                        (self.trUtf8(u"Website:"), self.projectdata["website"]),
                        (self.trUtf8(u"Apoio:"), self.projectdata["funding"])]

        self.datagroup = ((self.description, self.trUtf8(u"Descrição"), ""),
                          (self.site, self.trUtf8(u"Localização"), ""),
                          (self.address, self.trUtf8(u"Contato"), ""))

    def reset_search(self):
        sheet = self.active_mdi_child()
        rows = sheet.rowCount()
        cols = sheet.columnCount()
        for row in range(rows):
            for col in range(cols):
                sheet.item(row, col).setBackground(QtGui.QColor(255,255,255))

    def get_filter(self, sheet):
        reply = QtGui.QMessageBox.question(self, self.trUtf8(u"Confirmação"),
                                           self.trUtf8(u"Deseja filtrar os registros?"),
                                           QtGui.QMessageBox.Yes, QtGui.QMessageBox.No)
        if reply == QtGui.QMessageBox.Yes:
            filter_expr = self.qbe(sheet, False)
        else:
            filter_expr = ""
        return filter_expr

    def qbe(self, tableArea, SQLSelect):
        query = QueryDialog(self.trUtf8(u"Filtro"),
                            tableName=[os.path.splitext(tableArea.filename)[0]],
                            fieldList=tableArea.headers,
                            fieldTypes=tableArea.types,
                            alias=False,
                            SQLSelectClause=SQLSelect,
                            icon=QtGui.QIcon(":/ecolog.png"),
                            parent=self)
        if query.exec_():
            expr = query.SQL
        else:
            expr = ""
        return expr


def main():
    app = QtGui.QApplication(sys.argv)
    locale = QtCore.QLocale.system().name()
    appTranslator = QtCore.QTranslator()
    if appTranslator.load("ECOLOG_" + locale, ":/"):
        app.installTranslator(appTranslator)
    QtCore.QSettings.setDefaultFormat(QtCore.QSettings.IniFormat)
    app.setOrganizationName("Ecoinformatics Studio")
    app.setOrganizationDomain("ecolog.sourceforge.net")
    app.setApplicationName(app.translate("global", "ECOLOG"))
    app.setWindowIcon(QtGui.QIcon(":/ecolog.png"))
    splash_pix = QtGui.QPixmap(":/splash.png")
    splash = QtGui.QSplashScreen(splash_pix, QtCore.Qt.WindowStaysOnTopHint)
    splash.setMask(splash_pix.mask())
    splash.show()
    app.processEvents()
    time.sleep(2)
    form = MainWindow()
    form.show()
    splash.finish(form)
    app.exec_()


if __name__ == '__main__':
    main()

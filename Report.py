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
#    NumPy 1.4+ (www.numpy.org)                                                  #
#    Matplotlib 0.98+ (matplotlib.org)                                           #
#    geopy 0.94+ (github.com/geopy/geopy)                                        #
#    scipy-cluster 0.20+ (code.google.com/archive/p/scipy-cluster)               #
#    fuzzywuzzy 0.10+ (github.com/seatgeek/fuzzywuzzy)                           #
#================================================================================#

from __future__ import division
import warnings
import matplotlib
matplotlib.use("Agg")
import matplotlib.pyplot as plt
import os, time, sys, math, locale, codecs, sqlite3
from hcluster import pdist, linkage, dendrogram, cophenet, squareform
import numpy
from fuzzywuzzy import fuzz
from fuzzywuzzy import process
from geopy.geocoders import GoogleV3
from collections import defaultdict
from htmlentitydefs import codepoint2name
from PyQt4 import QtCore, QtGui

from Webservices import (checkCoL, searchCoL)
from Useful import (degtodms, extenso, find, get_unit, htmlescape, iif, is_online, parse_name, 
                    percent, quote_identifier, remove_duplicates, roman, to_float, to_int, 
                    truncate, unicode_to_ascii)
from Calc import (biodiv, ca, cca, morisita_horn, ochiai, pca, pcoa, rarefact, rda, sample_diversity)
from NMDS import nmds

SAMPLE = 0
INDIVIDUAL = 1
FAMILY = 2
SPECIES = 3
COLLECTOR = 4
NUMCOL = 5
DATE = 6
OBS = 7
LOCALITY = 8
LATITUDE = 9
LONGITUDE = 10
ELEVATION = 11

#--- Disable all warnings
warnings.filterwarnings("ignore")

class DNA:
    # From Patrick O'Brien, "Beginning Python for Bioinformatics"
    # http://www.onlamp.com/pub/a/python/2002/10/17/biopython.html
    
    """Class representing DNA as a string sequence."""
    basecomplement = {'A': 'T', 'C': 'G', 'T': 'A', 'G': 'C'}
        
    def __init__(self, s):
        """Create DNA instance initialized to string s."""
        self.seq = s
        
    def transcribe(self):
        """Return as rna string."""
        return self.seq.replace('T', 'U')
        
    def reverse(self):
        """Return dna string in reverse order."""
        letters = list(self.seq)
        letters.reverse()
        return ''.join(letters)
        
    def complement(self):
        """Return the complementary dna string."""
        letters = list(self.seq)
        letters = [self.basecomplement[base] for base in letters]
        return ''.join(letters)
        
    def reversecomplement(self):
        """Return the reverse complement of the dna string."""
        letters = list(self.seq)
        letters.reverse()
        letters = [self.basecomplement[base] for base in letters]
        return ''.join(letters)
        
    def gc(self):
        """Return the percentage of dna composed of G+C."""
        s = self.seq
        gc = s.count('G') + s.count('C')
        return gc * 100.0 / len(s)
        
    def codons(self):
        """Return list of codons for the dna string."""
        s = self.seq
        end = len(s) - (len(s) % 3) - 1
        codons = [s[i:i+3] for i in range(0, end, 3)]
        return codons
        
    def frequency(self, cds):
        """Return frequency table of codons for the dna string."""
        tally = defaultdict(int)
        for x in cds:
            tally[x] += 1
        return tally.items()

class Report(QtGui.QTextBrowser):
    def __init__(self, filename, projectdata, headers, types, data, filter):
        super(QtGui.QTextBrowser, self).__init__()
        self.setAttribute(QtCore.Qt.WA_DeleteOnClose)
        self.filename = filename
        self.title = QtCore.QFileInfo(self.filename).fileName()
        self.projectdata = projectdata
        self.headers = headers
        self.types = types
        self.data = data
        self.filter = filter
        self.setMinimumSize(500, 300)
        self.setFont(QtGui.QFont("Courier", 9))
        self.setWindowTitle(self.title)
        
    def Header(self, outfile, title):
        projectname = self.projectdata["title"]
        surveytype = self.projectdata["method"]
        header = "*****  ECOLOG  *****"
        outfile.write("<!DOCTYPE HTML PUBLIC ""-//W3C//DTD HTML 3.2//EN"">")
        outfile.write("<html>\n")
        outfile.write("<head>\n")
        outfile.write("<title>" + projectname + "</title>\n")
        outfile.write("</head>\n")
        outfile.write("<body>\n")
        outfile.write("<pre>\n")
        outfile.write(' '.join('*'*len(header)) + "\n")
        outfile.write(' '.join(header) + "\n")
        outfile.write(' '.join('*'*len(header)) + "\n")
        outfile.write("</pre>\n")	
        outfile.write(self.trUtf8(u"Data - ") + time.strftime("%d/%m/%Y", time.localtime()) + "<br>\n")
        outfile.write(self.trUtf8(u"Hora - ") + time.strftime("%H:%M:%S", time.localtime()) + "<br><br>\n\n")
        outfile.write(self.trUtf8(u"Projeto: ") + projectname + "<br><br>\n")
        outfile.write(self.trUtf8(u"Tipo de levantamento: ") + surveytype + "<br><br>\n\n")
        outfile.write(title + "<br>\n")
        if len(self.filter) > 0:
            pos = self.filter.rfind("WHERE")
            if pos > 0: self.filter = self.filter[pos + 6:]
            outfile.write("<br>\n" + self.trUtf8(u"FILTRO: ") + self.filter + "<br><br>\n")
            
    def Footer(self, outfile, count, total):
        if count > 0 or total > 0:
            outfile.write("<br>--<br>\n")
            outfile.write(self.trUtf8(u"TOTAL DE REGISTROS ARMAZENADOS: ") + str(total) + "<br>\n")
            outfile.write(self.trUtf8(u"NÚMERO DE REGISTROS RECUPERADOS: ") + str(count) + "<br>\n")
            outfile.write(self.trUtf8(u"% DE RECUPERAÇÃO/TOTAL: ") + "{:.2f}".format(percent(count, total)) + "\n")
        outfile.write("</body>\n")
        outfile.write("</html>\n")
        
    def Scissor(self, outfile):
        outfile.write("<p><img src="'scissor.png'" align=""Bottom"" alt="" "">" + '-'*160 + "</p>\n")
        
    def Geral(self):
        sqlCreateStr = "CREATE TABLE Temp("
        for i in range(len(self.headers)):
            if self.types[i] == "text":
                sqlCreateStr += quote_identifier(self.headers[i], "replace") + " TEXT"
            elif self.types[i] == "numeric":
                sqlCreateStr += quote_identifier(self.headers[i], "replace") + " NUMERIC"
            if i < len(self.headers)-1: sqlCreateStr += ", "
        sqlCreateStr = sqlCreateStr + ")"

        sqlInsertStr = "INSERT INTO Temp VALUES("
        for i in range(len(self.headers)):
            sqlInsertStr += "?"
            if i < len(self.headers)-1: sqlInsertStr += ", "
        sqlInsertStr = sqlInsertStr + ")"
        
        sqlSelectStr = "SELECT DISTINCT TRIM(" + self.headers[FAMILY] + "), TRIM(" + self.headers[SPECIES] + ") FROM Temp" 
        if len(self.filter) > 0: 
            sqlSelectStr += self.filter
        sqlSelectStr += " ORDER BY TRIM(" + self.headers[FAMILY] + "), TRIM(" + self.headers[SPECIES] + ")"

        db_data = tuple(self.data)
        db_connection = sqlite3.connect(":memory:")
        db_cursor = db_connection.cursor()
        db_cursor.execute(sqlCreateStr)
        db_cursor.executemany(sqlInsertStr, db_data)
        db_cursor.execute(sqlSelectStr)
        all_rows = db_cursor.fetchall()
        db_cursor.close()
        db_connection.close()
        
        total = len(all_rows)
        count = 0
        grpCheck = ""
        grpEval = ""
        
        outfile = open(unicode(self.filename), 'w')
        self.Header(outfile, self.trUtf8(u"RELATÓRIO GERAL DE ESPÉCIES"))

        for row in all_rows:
            grpCheck = str(row[0])
            if grpEval <> grpCheck:
                grpEval = grpCheck
                outfile.write("<br>" + grpCheck.upper() + "<br>\n")
            genus, cf, species, author1, subsp, infraname, author2 = parse_name(str(row[1]))
            outfile.write("&nbsp;&nbsp;&nbsp;&nbsp;<i>" + genus + " " + cf + " " + species + "</i> " + \
                author1 + " " + subsp + " <i>" + infraname + "</i> " + author2 + "<br>\n")
            count += 1
    
        self.Footer(outfile, count, total)
        outfile.close()
        self.setSource(QtCore.QUrl(self.title))
            
    def Checklist(self):
        sqlCreateStr = "CREATE TABLE Temp("
        for i in range(len(self.headers)):
            if self.types[i] == "text":
                sqlCreateStr += quote_identifier(self.headers[i], "replace") + " TEXT"
            elif self.types[i] == "numeric":
                sqlCreateStr += quote_identifier(self.headers[i], "replace") + " NUMERIC"
            if i < len(self.headers)-1: sqlCreateStr += ", "
        sqlCreateStr = sqlCreateStr + ")"

        sqlInsertStr = "INSERT INTO Temp VALUES("
        for i in range(len(self.headers)):
            sqlInsertStr += "?"
            if i < len(self.headers)-1: sqlInsertStr += ", "
        sqlInsertStr = sqlInsertStr + ")"
        
        sqlSelectStr = "SELECT " + self.headers[INDIVIDUAL] + ", TRIM(" + self.headers[FAMILY] + "), TRIM(" + \
                self.headers[SPECIES] + "), TRIM(" + self.headers[COLLECTOR] + "), " + self.headers[NUMCOL] + ", " + \
                self.headers[DATE] + ", TRIM(" + self.headers[LOCALITY] + "), " + \
                quote_identifier(self.headers[ELEVATION], "ignore") + " FROM Temp"
        if len(self.filter) > 0:
            sqlSelectStr += self.filter
        sqlSelectStr += " ORDER BY TRIM(" + self.headers[FAMILY] + "), TRIM(" + self.headers[SPECIES] + "), TRIM(" + \
                self.headers[LOCALITY] + "), TRIM(" + self.headers[COLLECTOR] + "), " + self.headers[DATE]
        
        db_data = tuple(self.data)
        db_connection = sqlite3.connect(":memory:")
        db_cursor = db_connection.cursor()
        db_cursor.execute(sqlCreateStr)
        db_cursor.executemany(sqlInsertStr, db_data)
        db_cursor.execute(sqlSelectStr)
        all_rows = db_cursor.fetchall()
        db_cursor.close()
        db_connection.close()
        
        mplace = self.projectdata["province"]
        if mplace is None: mplace = ""
        mloc = self.projectdata["locality"]
        if mloc is None: mloc = ""
        
        total = len(all_rows)
        count = 0
        grpCheck = ""
        grpEval = ""
        subCheck = ""
        subEval = ""

        outfile = open(unicode(self.filename), 'w')
        self.Header(outfile, self.trUtf8(u"CATÁLOGO DE COLETAS"))
        
        outfile.write("<dl>")
        for row in all_rows:
            grpCheck = str(row[1])
            if grpEval <> grpCheck:
                grpEval = grpCheck
                outfile.write("<br>" + grpCheck + "<br>\n")

            subCheck = str(row[2])
            if subEval <> subCheck:
                subEval = subCheck
                genus, cf, species, author1, subsp, infraname, author2 = parse_name(str(row[2]))
                outfile.write("<dt><i>" + genus + " " + cf + " " + species + "</i> " + \
                author1 + " " + subsp + " <i>" + infraname + "</i> " + author2 + "</dt>\n")
            
            line = "<dd>"
            line = "  Mun. " + mplace.strip() + ": " + mloc.strip() + ", " + str(row[6]) + ". "
            
            if row[7] is not None:
                if '-' in str(row[7]):
                    alt = row[7]
                else:
                    alt = to_int(row[7])
            else:	
                alt = 0
            if alt > 0:
                line += self.trUtf8(" Alt. ") + str(alt)
            elif alt < 0:
                line += self.trUtf8(" Prof. ") + str(alt * -1)
            line += ' ' + get_unit(self.headers[ELEVATION]) + ". "
            
            line += "<u>" + iif(row[3] != None, str(row[3]), "") + " " + \
                iif(row[4] != None, str(to_int(row[4])), "s/n") + "</u>. " + \
                iif(row[5] != None, str(row[5]).replace('-', '.'), "")
            outfile.write(line + "</dd><br>\n")
            count += 1
        outfile.write("</dl>")

        self.Footer(outfile, count, total)
        outfile.close()
        self.setSource(QtCore.QUrl(self.title))
        
    def Label(self, formato):
        mesext = [self.trUtf8("Janeiro"), self.trUtf8("Fevereiro"), self.trUtf8("Março"), self.trUtf8("Abril"), 
            self.trUtf8("Maio"), self.trUtf8("Junho"), self.trUtf8("Julho"), self.trUtf8("Agosto"),
            self.trUtf8("Setembro"), self.trUtf8("Outubro"), self.trUtf8("Novembro"), self.trUtf8("Dezembro")]
            
        sqlCreateStr = "CREATE TABLE Temp("
        for i in range(len(self.headers)):
            if self.types[i] == "text":
                sqlCreateStr += quote_identifier(self.headers[i], "replace") + " TEXT"
            elif self.types[i] == "numeric":
                sqlCreateStr += quote_identifier(self.headers[i], "replace") + " NUMERIC"
            if i < len(self.headers)-1: sqlCreateStr += ", "
        sqlCreateStr = sqlCreateStr + ")"

        sqlInsertStr = "INSERT INTO Temp VALUES("
        for i in range(len(self.headers)):
            sqlInsertStr += "?"
            if i < len(self.headers)-1: sqlInsertStr += ", "
        sqlInsertStr = sqlInsertStr + ")"
        
        sqlSelectStr = "SELECT " + self.headers[INDIVIDUAL] + ", TRIM(" + self.headers[FAMILY] + "), TRIM(" + \
                self.headers[SPECIES] + "), TRIM(" + self.headers[COLLECTOR] + "), " + self.headers[NUMCOL] + ", " + \
                self.headers[DATE] + ", TRIM(" + self.headers[OBS] + "), TRIM(" + self.headers[LOCALITY] + "), " + \
                quote_identifier(self.headers[ELEVATION], "ignore") + ", " + \
                self.headers[LATITUDE] + ", " + self.headers[LONGITUDE] + " FROM Temp"
        if len(self.filter) > 0:
            sqlSelectStr += self.filter
        sqlSelectStr += " ORDER BY TRIM(" + self.headers[FAMILY] + "), TRIM(" + self.headers[SPECIES] + "), TRIM(" + \
                self.headers[LOCALITY] + "), TRIM(" + self.headers[COLLECTOR] + "), " + self.headers[DATE]
    
        db_data = tuple(self.data)
        db_connection = sqlite3.connect(":memory:")
        db_cursor = db_connection.cursor()
        db_cursor.execute(sqlCreateStr)
        db_cursor.executemany(sqlInsertStr, db_data)
        db_cursor.execute(sqlSelectStr)
        all_rows = db_cursor.fetchall()
        db_cursor.close()
        db_connection.close()

        l_institute = self.projectdata["institution"].encode("utf-8")
        if l_institute is None: l_institute = ""
        l_header = self.projectdata["title"].encode("utf-8")
        if l_header is None: l_header = ""
        l_footer = self.projectdata["funding"].encode("utf-8")
        if l_footer is None: l_footer = ""
        mcountry = self.projectdata["country"].encode("utf-8")
        if mcountry is None: mcountry = ""
        mstate = self.projectdata["state"].encode("utf-8")
        if mstate is None: mstate = ""
        mprov = self.projectdata["province"].encode("utf-8")
        if mprov is None: mprov = ""
        mloc = self.projectdata["locality"].encode("utf-8")
        if mloc is None: mloc = ""

        outfile = open(unicode(self.filename), 'w')
        outfile.write("<html>\n")
        outfile.write("<head>\n")
        outfile.write("<title>" + l_header + "</title>\n")
        outfile.write("</head>\n")
        outfile.write("<body>\n")
        
        for row in all_rows:
            mproc = mcountry.upper() + ", " + mstate.strip() + ", Mun. " + mprov.strip() + ", " + mloc.strip()
            if row[7] is not None:
                mproc += ", " + str(row[7]) + "."

            if row[9] is not None or row[10] is not None:
                vlatdeg, vlatmin, vlatsec = degtodms(to_float(row[9]))
                vlondeg, vlonmin, vlonsec = degtodms(to_float(row[10]))
                if vlatdeg < 0.0:
                    vlath = "S"
                else:
                    vlath = "N"
                if vlondeg < 0.0:
                    vlonh = "W"
                else:
                    vlonh = "E"
                mproc += " (Coord. " + str(vlatdeg) + "<sup>o</sup>" + str(vlatmin) + "'" + str(vlatsec) + "'' " + vlath \
                    + ", " + str(vlondeg) + "<sup>o</sup>" + str(vlonmin) + "'" + str(vlonsec) + "'' " + vlonh + "). "

            if row[8] is not None:
                if '-' in str(row[8]):
                    alt = row[8]
                else:
                    alt = to_int(row[8])
            else:
                alt = 0
            if alt > 0:
                mproc += self.trUtf8(" Alt. ") + str(alt)
            elif alt < 0:
                mproc += self.trUtf8(" Prof. ") + str(alt * -1)
            mproc += ' ' + get_unit(self.headers[ELEVATION]) + ". "
        
            mfamily = str(row[1])
            mgenus, mcf, mspecies, mauthor1, msubsp, minfraname, mauthor2 = parse_name(str(row[2]))
            mgenus = "<i>" + mgenus + "</i>"
            mspecies = mgenus + " " + cf + " <i>" + mspecies + "</i> " + mauthor1
            if len(msubsp) > 0:
                mspecies += " " + msubsp + " <i>" + minfraname + "</i> " + mauthor2
        
            if row[3] is not None:
                mcoletor = "Col. "  + str(row[3]) + " "
            else:
                mcoletor = ""
            
            if row[4] is not None:
                mcoletor += str(to_int(row[4]))
            else:
                mcoletor += "s/n"

            mdatacol = self.trUtf8("Data: ")
            if formato == 1:
                mdatacol += str(row[5])
            elif formato == 2:
                mdatacol += roman(str(row[5]))
            elif formato == 3:
                mdatacol += extenso(str(row[5]), mesext)
                
            if row[6] is not None:
                mobs = str(row[6])
                if not mobs.endswith('.'): mobs += '.'
            else:
                mobs = ""
        
            outfile.write("\n<center>\n")
            outfile.write("<table border=0 cellspacing=2 cellpadding=2 width=""80%"">\n")
            outfile.write("<tr align=""Center""><td align=""Center"" colspan=2>" + htmlescape(l_institute) + "</td></tr>\n")
            outfile.write("<tr align=""Center""><td align=""Center"" colspan=2>" + htmlescape(l_header) + "</td></tr>\n")
            outfile.write("<tr align=""Center""><td align=""Left"" colspan=2>" + mfamily + "</td></tr>\n")
            outfile.write("<tr align=""Center""><td align=""Left"" colspan=2>" + mspecies + "</td></tr>\n")
            outfile.write("<tr align=""Center""><td align=""Left"" colspan=2>" + mproc  + "</td></tr>\n")
            outfile.write("<tr align=""Center""><td align=""Left"" colspan=2>" + mobs + "</td></tr>\n")
            outfile.write("<tr align=""Center""><td align=""Left"" colspan=2>" + mcoletor + "</td></tr>\n")
            outfile.write("<tr align=""Center""><td align=""Left"" colspan=2>" + mdatacol + "</td></tr>\n")
            outfile.write("<tr align=""Center""><td align=""Left"" colspan=2>" + "Det." + "</td></tr>\n")
            outfile.write("<tr align=""Center""><td align=""Left"" colspan=2>" + self.trUtf8("Data:") + "</td></tr>\n")
            outfile.write("<tr align=""Center""><td align=""Center"" colspan=2>" + htmlescape(l_footer) + "</td></tr>\n")
            outfile.write("</table>\n")
            outfile.write("</center>\n")
            self.Scissor(outfile)
            
        outfile.write("</body>\n")
        outfile.write("</html>\n")
        outfile.close()
        self.setSource(QtCore.QUrl(self.title))
        
    def Stats(self, option, graph_it):
        sqlCreateStr = "CREATE TABLE Temp("
        for i in range(len(self.headers)):
            if self.types[i] == "text":
                sqlCreateStr += quote_identifier(self.headers[i], "replace") + " TEXT"
            elif self.types[i] == "numeric":
                sqlCreateStr += quote_identifier(self.headers[i], "replace") + " NUMERIC"
            if i < len(self.headers)-1: sqlCreateStr += ", "
        sqlCreateStr = sqlCreateStr + ")"

        sqlInsertStr = "INSERT INTO Temp VALUES("
        for i in range(len(self.headers)):
            sqlInsertStr += "?"
            if i < len(self.headers)-1: sqlInsertStr += ", "
        sqlInsertStr = sqlInsertStr + ")"
        
        db_data = tuple(self.data)
        db_connection = sqlite3.connect(":memory:")
        db_cursor = db_connection.cursor()
        db_cursor.execute(sqlCreateStr)
        db_cursor.executemany(sqlInsertStr, db_data)
        
        outfile = open(unicode(self.filename), 'w')
        self.Header(outfile, self.trUtf8(u"RELATÓRIO ESTATÍSTICO"))
        
        if option == 1:
            outfile.write("<br>\n")
            outfile.write(self.trUtf8(u"Estatística de FAMÍLIAS"))
            outfile.write("<br><br>\n")
            outfile.write("<table border=1 cellspacing=1 cellpadding=1 width=""100%"">\n")
            outfile.write("<tr><th>" + self.trUtf8(u"FAMÍLIA") + "</th>")
            outfile.write("<th>" + self.trUtf8(u"NO.SPP.") + "</th>")
            outfile.write("<th>" + self.trUtf8(u"%SPP") + "</th>")
            outfile.write("</tr>\n")
            
            db_cursor.execute("SELECT COUNT(DISTINCT TRIM(" + self.headers[SPECIES] + ")) FROM Temp")
            s = int(db_cursor.fetchone()[0])
            
            f = 0
            vetfam = []
            db_cursor.execute("SELECT DISTINCT TRIM(" + self.headers[FAMILY] + ")" + \
                            " FROM Temp ORDER BY TRIM(" + self.headers[FAMILY] + ")")
            while True:
                row = db_cursor.fetchone()
                if row == None: break
                vfamilia = row[0]
                if vfamilia != None:
                    f += 1
                    vetfam.append(str(vfamilia))
            
            if graph_it:
                families = []
                frequency = []
                
            for i in range(f):
                db_cursor.execute("SELECT COUNT(DISTINCT TRIM(" + self.headers[SPECIES] + ")" + \
                            ") FROM Temp WHERE TRIM(" + self.headers[FAMILY] + ") = '" + str(vetfam[i]) + "'")
                vconta = int(db_cursor.fetchone()[0])
                outfile.write("<tr>")	
                outfile.write("<td align=""Left"">" + vetfam[i] + "</td>\n")
                outfile.write("<td align=""Center"">" + str(vconta) + self.trUtf8(u" espécie(s)") + "</td>\n")
                outfile.write("<td align=""Center"">" + "{:.1f}".format(percent(vconta, s)) + "</td>\n")
                outfile.write("</tr>")
                if graph_it:
                    families.append(vetfam[i])
                    frequency.append(vconta * 10.0)
            
            outfile.write("</table>\n\n")
            outfile.write("<br>\n<br>--<br>\n")
            outfile.write(self.trUtf8(u"TOTAL DE FAMÍLIAS = ") + str(f) + self.trUtf8(u" família(s)") + "<br>\n")
            outfile.write(self.trUtf8(u"TOTAL DE ESPÉCIES = ") + str(s) + self.trUtf8(u" espécie(s)") + "<br>\n")
            
            if graph_it:
                frequency, families = (list(t) for t in zip(*sorted(zip(frequency, families))))
                y_pos = numpy.arange(len(families))
                figf = os.path.splitext(os.path.basename(str(self.filename)))[0] + ".png"
                plt.clf()
                plt.barh(y_pos, frequency, align="center", alpha=0.4)
                plt.xlim([0, max(frequency) + 1])
                plt.ylim([y_pos.max() + 1, y_pos.min() - 1])
                plt.yticks(y_pos, families)
                plt.title(self.trUtf8(u"FREQUÊNCIA DE ESPÉCIES POR FAMÍLIA"))
                plt.ylabel(self.trUtf8(u"FAMÍLIAS"))
                plt.xlabel(self.trUtf8(u"FREQUÊNCIA RELATIVA (%)"))
                plt.grid(False)
                plt.tight_layout()
                plt.savefig(figf, dpi=72)
                outfile.write("<img src='" + figf + "'>\n")
                
        elif option == 2:
            outfile.write("<br>\n")
            outfile.write(self.trUtf8(u"Estatística de GÊNEROS"))
            outfile.write("<br><br>\n")
            outfile.write("<table border=1 cellspacing=1 cellpadding=1 width=""100%"">\n")
            outfile.write("<tr><th>" + self.trUtf8(u"GÊNERO") + "</th>")
            outfile.write("<th>" + self.trUtf8(u"NO.SPP.") + "</th>")
            outfile.write("<th>" + self.trUtf8(u"%SPP") + "</th>")
            outfile.write("</tr>\n")
            
            db_cursor.execute("SELECT COUNT(DISTINCT TRIM(" + self.headers[SPECIES] + ")) FROM Temp")
            s = int(db_cursor.fetchone()[0])
            
            db_cursor.execute("SELECT DISTINCT TRIM(" + self.headers[SPECIES] + ") FROM Temp")
            g = 0
            vetgen = []
            while True:
                row = db_cursor.fetchone()
                if row == None: break
                vgenero = str(row[0]).split()[0]
                if len(vgenero) > 0 and not vgenero in vetgen: 
                    g += 1
                    vetgen.append(str(vgenero))
            
            if graph_it:
                genera = []
                frequency = []
                
            for i in range(g):
                db_cursor.execute("SELECT COUNT(DISTINCT TRIM(" + self.headers[SPECIES] + ")) FROM Temp WHERE SUBSTR(TRIM(" + \
                                    self.headers[SPECIES] + "), 1, " + str(len(vetgen[i])) + ") = '" + \
                                    str(vetgen[i]) + "'")
                vconta = int(db_cursor.fetchone()[0])
                outfile.write("<tr>")	
                outfile.write("<td align=""Left""><i>" + vetgen[i] + "</i></td>\n")
                outfile.write("<td align=""Center"">" + str(vconta) + self.trUtf8(u" espécie(s)") + "</td>\n")
                outfile.write("<td align=""Center"">" + "{:.1f}".format(percent(vconta, s)) + "</td>\n")
                outfile.write("</tr>")
                if graph_it:
                    genera.append(vetgen[i])
                    frequency.append(vconta * 10.0)
                    
            outfile.write("</table>\n\n")
            outfile.write("<br>\n<br>--<br>\n")
            outfile.write(self.trUtf8(u"TOTAL DE GÊNEROS = ") + str(g) + self.trUtf8(u" gêneros(s)") + "<br>\n")
            outfile.write(self.trUtf8(u"TOTAL DE ESPÉCIES = ") + str(s) + self.trUtf8(u" espécie(s)") + "<br>\n")
            
            if graph_it:
                frequency, genera = (list(t) for t in zip(*sorted(zip(frequency, genera))))
                y_pos = numpy.arange(len(genera))
                figf = os.path.splitext(os.path.basename(str(self.filename)))[0] + ".png"
                plt.clf()
                plt.barh(y_pos, frequency, align="center", alpha=0.4)
                plt.xlim([0, max(frequency) + 1])
                plt.ylim([y_pos.max() + 1, y_pos.min() - 1])
                plt.yticks(y_pos, genera)
                plt.title(self.trUtf8(u"FREQUÊNCIA DE ESPÉCIES POR GÊNERO"))
                plt.ylabel(self.trUtf8(u"GÊNEROS"))
                plt.xlabel(self.trUtf8(u"FREQUÊNCIA RELATIVA (%)"))
                plt.grid(False)
                plt.tight_layout()
                ax = plt.subplot(111)
                for label in ax.get_yticklabels():
                    label.set_fontstyle("italic")
                plt.savefig(figf, dpi=72)
                outfile.write("<img src='" + figf + "'>\n")
        
        elif option == 3:
            outfile.write("<br>\n")
            outfile.write(self.trUtf8(u"Estatística de ESPÉCIES"))
            outfile.write("<br><br>\n")
            outfile.write("<table border=1 cellspacing=1 cellpadding=1 width=""100%"">\n")
            outfile.write("<tr><th>"+ self.trUtf8(u"ESPÉCIE") + "</th>")
            outfile.write("<th>" + self.trUtf8(u"NO.IND.") + "</th>")
            outfile.write("<th>" + self.trUtf8(u"DOM.%") + "</th>")
            outfile.write("<th>" + self.trUtf8(u"FREQ") + "</th>")
            outfile.write("<th>" + self.trUtf8(u"F.R.%") + "</th>")
            outfile.write("<th>" + self.trUtf8(u"CATEGORIA") + "</th>")
            outfile.write("</tr>\n")
            
            db_cursor.execute("SELECT DISTINCT TRIM(" + self.headers[SAMPLE] + ") FROM Temp")
            l = len(db_cursor.fetchall())
            
            db_cursor.execute("SELECT DISTINCT TRIM(" + self.headers[SPECIES] + ") FROM Temp")
            s = len(db_cursor.fetchall())
            
            db_cursor.execute("SELECT COUNT(*) FROM Temp")
            n = int(db_cursor.fetchone()[0]) 
            
            if graph_it:
                species = []
                frequency = []
                
            ni = []
            db_cursor.execute("SELECT DISTINCT TRIM(" + self.headers[SPECIES] + ")" + \
                            " FROM Temp ORDER BY TRIM(" + self.headers[SPECIES] + ")")
            while True:
                row = db_cursor.fetchone()
                if row == None: break
                vspp = str(row[0])
                vgenero, vcf, vespecie, vautor1, vsubsp, vinfranome, vautor2 = parse_name(vspp)
                vnome = vgenero + " " + vespecie + " " + iif(vinfranome != None, vinfranome, "")
                if graph_it:
                    vsigla = truncate(vgenero, 4, " ") + "." + truncate(vespecie, 4, " ")
                num_cursor = db_connection.cursor()
                num_cursor.execute("SELECT * FROM Temp WHERE " + self.headers[SPECIES] + " = '" + vspp + "'")
                found = num_cursor.fetchone() != None
                if found:
                    num_cursor.execute("SELECT COUNT(*), COUNT(DISTINCT TRIM(" + self.headers[SAMPLE] + ")" + \
                                ") FROM Temp WHERE TRIM(" + self.headers[SPECIES] + ") = '" + vspp + "'")
                    num_row = num_cursor.fetchone()
                    nind = int(num_row[0])
                    freq = int(num_row[1])
                    ni.append(float(nind))
                    c = percent(freq, l)
                    if c > 50.0:
                        cons = self.trUtf8(u"Constante")
                    elif c >= 25.0 and c <= 50.0:
                        cons = self.trUtf8(u"Acessória")
                    elif c < 25.0:
                        cons = self.trUtf8(u"Acidental")
                    outfile.write("<tr>")	
                    outfile.write("<td align=""Left""><i>" + vnome + "</i></td>\n")
                    outfile.write("<td align=""Center"">" + str(nind) + self.trUtf8(u" indivíduo(s)") + "</td>\n")
                    outfile.write("<td align=""Center"">" + "{:.1f}".format(percent(nind, n)) + "</td>\n")
                    outfile.write("<td align=""Center"">" + str(freq) + "</td>\n")
                    outfile.write("<td align=""Center"">" + "{:.1f}".format(c) + "</td>\n")
                    outfile.write("<td align=""Center"">" + cons + "</td>\n")
                    outfile.write("</tr>")
                    if graph_it:
                        species.append(vsigla)
                        frequency.append(c)
                num_cursor.close()
            outfile.write("</table>\n\n")
            outfile.write("<br>\n<br>--<br>\n")
            outfile.write(self.trUtf8(u"TOTAL DE ESPÉCIES = ") + str(s) + self.trUtf8(u" espécie(s)") + "<br>\n")
            outfile.write(self.trUtf8(u"TOTAL DE INDIVÍDUOS = ") + str(n) + self.trUtf8(u" indivíduo(s)") + "<br>\n\n")
            
            outfile.write("<br>" + self.trUtf8(u"Análise de DIVERSIDADE") + "<br>\n\n")
            d1, d2, c, h, pie, d, m, j = biodiv(s, n, ni)
            outfile.write("<br>" + self.trUtf8(u"RIQUEZA DE ESPÉCIES") + "<br>\n")
            outfile.write("<br>\n")
            outfile.write(self.trUtf8(u" Índice de Margalef  (D1)      = ") + "{:.5f}".format(d1) + "<br>\n")
            outfile.write(self.trUtf8(u" Índice de Menhinick (D2)      = ") + "{:.5f}".format(d2) + "<br>\n\n")
            outfile.write("<br>" + self.trUtf8(u"DIVERSIDADE") + "<br>\n")
            outfile.write("<br>\n")
            outfile.write(self.trUtf8(u" Índice de Simpson (C)         = ") + "{:.5f}".format(c) + "<br>\n")
            outfile.write(self.trUtf8(u" Índice de Shannon-Weaver (H') = ") + "{:.5f}".format(h) + "<br>\n")
            outfile.write(self.trUtf8(u" Índice de Hurlbert (PIE)      = ") + "{:.5f}".format(pie) + "<br>\n")
            outfile.write(self.trUtf8(u" Índice de Berger-Parker (d)   = ") + "{:.5f}".format(d) + "<br>\n")
            outfile.write(self.trUtf8(u" Índice de McIntosh (M)        = ") + "{:.5f}".format(m) + "<br>\n\n")
            outfile.write(u"<br>" + self.trUtf8(u"EQUITABILIDADE") + "<br>\n")
            outfile.write("<br>\n")
            outfile.write(self.trUtf8(u" Equitabilidade (J)            = ") + "{:.5f}".format(j) + "<br>\n")
            
            if graph_it:
                frequency, species = (list(t) for t in zip(*sorted(zip(frequency, species))))
                y_pos = numpy.arange(len(species))
                figf = os.path.splitext(os.path.basename(str(self.filename)))[0] + ".png"
                plt.clf()
                plt.barh(y_pos, frequency, align="center", alpha=0.4)
                plt.xlim([0, max(frequency) + 10])
                plt.ylim([y_pos.max() + 1, y_pos.min() - 1])
                plt.yticks(y_pos, species)
                plt.title(self.trUtf8(u"FREQUÊNCIA DE INDIVÍDUOS POR ESPÉCIE"))
                plt.ylabel(self.trUtf8(u"ESPÉCIE"))
                plt.xlabel(self.trUtf8(u"FREQUÊNCIA RELATIVA (%)"))
                plt.grid(False)
                plt.tight_layout()
                plt.savefig(figf, dpi=72)
                outfile.write("<img src='" + figf + "'>\n")
        
        elif option == 4:
            outfile.write("<br>\n")
            outfile.write(self.trUtf8(u"Estatística de LOCAIS"))
            outfile.write("<br><br>\n")
            outfile.write("<table border=1 cellspacing=1 cellpadding=1 width=""100%"">\n")
            outfile.write("<tr><th>" + self.trUtf8(u"LOCAL") + "</th>")
            outfile.write("<th>" + self.trUtf8(u"NO.SPP.") + "</th>")
            outfile.write("<th>" + self.trUtf8(u"NO.IND.") + "</th>")
            outfile.write("<th>" + self.trUtf8(u"%SPP") + "</th>")
            outfile.write("</tr>\n")
            
            db_cursor.execute("SELECT DISTINCT TRIM(" + self.headers[SAMPLE] + ") FROM Temp")
            l = len(db_cursor.fetchall())
            
            db_cursor.execute("SELECT DISTINCT TRIM(" + self.headers[SPECIES] + ") FROM Temp")
            s = len(db_cursor.fetchall())
            
            db_cursor.execute("SELECT DISTINCT TRIM(" + self.headers[SAMPLE] + ") FROM Temp ORDER BY TRIM(" + 
                    self.headers[SAMPLE] + ")")
            
            if graph_it:
                x = [0]
                y = [0]
                i = 1
            
            while True:
                row = db_cursor.fetchone()
                if row == None: break
                vloc = str(row[0])
                aux_cursor = db_connection.cursor()
                aux_cursor.execute("SELECT COUNT(DISTINCT TRIM(" + self.headers[SPECIES] + ")), COUNT(TRIM(" + \
                            self.headers[SPECIES] + ")) FROM Temp WHERE TRIM(" + self.headers[SAMPLE] + ") = '" + vloc + "'")
                aux_row = aux_cursor.fetchone()
                nspp = int(aux_row[0])
                nind = int(aux_row[1])
                outfile.write("<tr>")	
                outfile.write("<td align=""Left"">" + vloc + "</td>\n")
                outfile.write("<td align=""Center"">" + str(nspp) + self.trUtf8(u" espécie(s)") + "</td>\n")
                outfile.write("<td align=""Center"">" + str(nind) + self.trUtf8(u" indivíduos(s)") + "</td>\n")
                outfile.write("<td align=""Center"">" + "{:.1f}".format(percent(nspp, s)) + "</td>\n")
                outfile.write("</tr>")
                aux_cursor.close()
                if graph_it:
                    x.append(i)
                    y.append(nspp)
                    i += 1
            outfile.write("</table>\n\n")	
            outfile.write("<br>\n<br>--<br>\n")
            outfile.write(self.trUtf8(u"TOTAL DE LOCAIS = ") + str(l) + self.trUtf8(u" locais") + "<br>\n")
            outfile.write(self.trUtf8(u"TOTAL DE ESPÉCIES = ") + str(s) + self.trUtf8(u" espécie(s)") + "<br>\n")
            
            if graph_it:
                yc = []
                for i in range(1, len(y)):
                    yc.append(y[i] + y[i - 1])
                yc.insert(0, y[0])
                yc.sort()
                figf = os.path.splitext(os.path.basename(str(self.filename)))[0] + ".png"
                plt.clf()
                plt.plot(x, yc, color="green", linewidth=3.0, alpha=0.6)
                plt.xlim([0, max(x) + 1])
                plt.ylim([0, max(yc) + 1])
                plt.title(self.trUtf8(u"CURVA DO COLETOR"))
                plt.xlabel(self.trUtf8(u"AMOSTRAS"))
                plt.ylabel(self.trUtf8(u"NÚMERO CUMULATIVO DE ESPÉCIES"))
                plt.grid(False)
                plt.tight_layout()
                plt.savefig(figf, dpi=72)
                outfile.write("<img src='" + figf + "'>\n")
        
        elif option == 5:
            outfile.write("<br>\n")
            outfile.write(self.trUtf8(u"Estatística de DESCRITORES"))
            outfile.write("<br><br>\n")
            outfile.write("<table border=1 cellspacing=1 cellpadding=1 width=""100%"">\n")
            outfile.write("<tr><th>" + self.trUtf8(u"NÚMERO") + "</th>")
            outfile.write("<th>" + self.trUtf8(u"DESCRITOR") + "</th>")
            outfile.write("<th>" + self.trUtf8(u"TIPO") + "</th>")
            outfile.write("<th>" + self.trUtf8(u"UNIDADE") + "</th>")
            outfile.write("<th>" + self.trUtf8(u"MÉDIA")  + "</th>")
            outfile.write("<th>" + self.trUtf8(u"DESVIO-PADRÃO") + "</th>")
            outfile.write("<th>" + self.trUtf8(u"MÍNIMO") + "</th>")
            outfile.write("<th>" + self.trUtf8(u"MÁXIMO") + "</th>")
            outfile.write("<th>" + self.trUtf8(u"CASOS") + "</th></tr>\n")
            
            descriptors = self.headers[12:len(self.headers)]
            datavalues = []
            values = []
            for j in range(12,len(self.headers)):
                values = [item[j] for item in self.data]
                datavalues.append(values)
            datavalues = zip(*datavalues)
            datatypes = []
            na = []
            mean = []
            sd = []
            minimum = []
            maximum = []
            for j in range(len(descriptors)):
                nc = 0
                data = []
                for i in range(len(datavalues)):
                    val = datavalues[i][j]
                    if val is None:
                        nc += 1
                    elif val.find('+') != -1:
                        val = eval(val)
                    data.append(to_float(val))
                    try:
                        val = to_float(val)
                        datatypes.append(self.trUtf8(u"Numérico"))
                    except ValueError:
                        datatypes.append(self.trUtf8(u"Texto"))
                minimum.append(min(data))
                maximum.append(max(data))
                mean.append(numpy.average(data))
                sd.append(numpy.std(data))
                na.append(nc)

            count = 0
            for i in range(len(descriptors)):
                descriptor = htmlescape(descriptors[i].encode("utf-8").split()[0])
                outfile.write("<tr>")
                outfile.write("<td align=""Left"">" + str(count + 1) + "</td>\n")
                outfile.write("<td align=""Left"">" + descriptor.capitalize() + "</td>\n")
                outfile.write("<td align=""Left"">" + datatypes[i] + "</td>\n")
                outfile.write("<td align=""Center"">" + get_unit(descriptors[i]) + "</td>\n")
                outfile.write("<td align=""Right"">" + "{:.3f}".format(mean[i]) + "</td>\n")
                outfile.write("<td align=""Right"">" + "{:.3f}".format(sd[i]) + "</td>\n")
                outfile.write("<td align=""Right"">" + "{:.2f}".format(minimum[i]) + "</td>\n")
                outfile.write("<td align=""Right"">" + "{:.2f}".format(maximum[i]) + "</td>\n")
                outfile.write("<td align=""Right"">" + str(len(self.data) - na[i]) + "</td>\n")
                outfile.write("</tr>")
                count += 1
            outfile.write("</table>\n\n")
            outfile.write("<br>\n<br>--<br>\n")
            outfile.write(self.trUtf8(u"TOTAL DE DESCRITORES = ") + str(count) + self.trUtf8(u" descritor(es)") + "<br>\n")
        
        elif option == 6:
            outfile.write("<br>\n")
            outfile.write(self.trUtf8(u"Estatística de VARIÁVEIS AMBIENTAIS"))
            outfile.write("<br><br>\n")
            outfile.write("<table border=1 cellspacing=1 cellpadding=1 width=""100%"">\n")
            outfile.write("<tr><th>" + self.trUtf8(u"NÚMERO") + "</th>")
            outfile.write("<th>" + self.trUtf8(u"VARIÁVEL") + "</th>")
            outfile.write("<th>" + self.trUtf8(u"TIPO") + "</th>")
            outfile.write("<th>" + self.trUtf8(u"UNIDADE") + "</th>")
            outfile.write("<th>" + self.trUtf8(u"MÉDIA")  + "</th>")
            outfile.write("<th>" + self.trUtf8(u"DESV.PAD.") + "</th>")
            outfile.write("<th>" + self.trUtf8(u"COEF.VAR. (.%)") + "</th>")
            outfile.write("<th>" + self.trUtf8(u"MÍN.") + "</th>")
            outfile.write("<th>" + self.trUtf8(u"MÁX.") + "</th>")
            outfile.write("<th>" + self.trUtf8(u"N") + "</th></tr>\n")
            
            variables = self.headers[4:len(self.headers)]
            datavalues = []
            values = []
            for j in range(4,len(self.headers)):
                values = [item[j] for item in self.data]
                datavalues.append(values)
            datavalues = zip(*datavalues)
            datatypes = []
            na = []
            mean = []
            sd = []
            cv = []
            minimum = []
            maximum = []
            for j in range(len(variables)):
                nc = 0
                data = []
                for i in range(len(datavalues)):
                    val = datavalues[i][j]
                    if val is None:
                        nc += 1
                    data.append(to_float(val))
                    try:
                        val = to_float(val)
                        datatypes.append(self.trUtf8(u"Numérico"))
                    except ValueError:
                        datatypes.append(self.trUtf8(u"Texto"))
                minimum.append(min(data))
                maximum.append(max(data))
                mean.append(numpy.average(data))
                sd.append(numpy.std(data))
                cv.append(numpy.std(data) / numpy.average(data))
                na.append(nc)

            count = 0
            for i in range(len(variables)):
                variable = htmlescape(variables[i].encode("utf-8").split()[0])
                outfile.write("<tr>")
                outfile.write("<td align=""Left"">" + str(count + 1) + "</td>\n")
                outfile.write("<td align=""Left"">" + variable.capitalize() + "</td>\n")
                outfile.write("<td align=""Left"">" + datatypes[i] + "</td>\n")
                outfile.write("<td align=""Center"">" + get_unit(variables[i]) + "</td>\n")
                outfile.write("<td align=""Right"">" + "{:.3f}".format(mean[i]) + "</td>\n")
                outfile.write("<td align=""Right"">" + "{:.3f}".format(sd[i]) + "</td>\n")
                outfile.write("<TD ALIGN=""Right"">" + "{:.3f}".format(cv[i] * 100.0) + "</td>\n")
                outfile.write("<td align=""Right"">" + "{:.2f}".format(minimum[i]) + "</td>\n")
                outfile.write("<td align=""Right"">" + "{:.1f}".format(maximum[i]) + "</td>\n")
                outfile.write("<td align=""Right"">" + str(len(self.data) - na[i]) + "</td>\n")
                outfile.write("</tr>")
                count += 1
            outfile.write("</table>\n\n")
            outfile.write("<br>\n<br>--<br>\n")
            outfile.write(self.trUtf8(u"TOTAL DE VARIÁVEIS = ") + str(count) + self.trUtf8(u" variáveis(es)") + "<br>\n")
        
        elif option == 7:
            outfile.write("<br>\n")
            outfile.write(self.trUtf8(u"Estatística de SEQUÊNCIAS"))
            outfile.write("<br><br>\n\n")
            outfile.write("<br>" + self.trUtf8(u"Composição de NUCLEOTÍDEOS") + "<br>\n")
            outfile.write("<table border=1 cellspacing=1 cellpadding=1 width=""100%"">\n")
            outfile.write("<tr><th>" + self.trUtf8(u"INDIVÍDUO") + "</th>")
            outfile.write("<th>A</th>")
            outfile.write("<th>%</th>")
            outfile.write("<th>T</th>")
            outfile.write("<th>%</th>")
            outfile.write("<th>C</th>")
            outfile.write("<th>%</th>")
            outfile.write("<th>G</th>")
            outfile.write("<th>%</th>")
            outfile.write("<th>Total</th>")
            outfile.write("<th>%CG</th>")
            outfile.write("</tr>\n")
            
            s = self.trUtf8(u"Sequência")
            seq = [s for s in self.headers if "SEQ" in s.upper()]
            pos = find(self.headers, seq[0])
            nseq = 0
            
            db_cursor.execute("SELECT " + self.headers[INDIVIDUAL] + ", " +  self.headers[pos] + \
                            " FROM Temp WHERE " + self.headers[pos] + " IS NOT NULL" + \
                            " ORDER BY " + self.headers[INDIVIDUAL])
            while True:
                row = db_cursor.fetchone()
                if row == None: break
                seq = str(row[1])
                if len(seq) > 0:
                    nseq += 1
                    seqlen = len(seq)
                    countA = seq.count('A')
                    countT = seq.count('T')
                    countC = seq.count('C')
                    countG = seq.count('G')
                    nuc = DNA(seq)
                    outfile.write("<tr>")	
                    outfile.write("<td align=""Left"">" + str(row[0]) + "</td>\n")
                    outfile.write("<td align=""Center"">" + str(countA) + "</td>\n")
                    outfile.write("<td align=""Center"">" + "{:.1f}".format((countA * 100.0 / seqlen)) + "</td>\n")
                    outfile.write("<td align=""Center"">" + str(countT) + "</TD>\n")
                    outfile.write("<td align=""Center"">" + "{:.1f}".format((countT * 100.0 / seqlen)) + "</td>\n")
                    outfile.write("<td align=""Center"">" + str(countC) + "</TD>\n")
                    outfile.write("<td align=""Center"">" + "{:.1f}".format((countC * 100.0 / seqlen)) + "</td>\n")
                    outfile.write("<td align=""Center"">" + str(countG) + "</TD>\n")
                    outfile.write("<td align=""Center"">" + "{:.1f}".format((countG * 100.0 / seqlen)) + "</td>\n")
                    outfile.write("<td align=""Center"">" + str(seqlen) + "</TD>\n")
                    outfile.write("<td align=""Center"">" + "{:.1f}".format(DNA.gc(nuc)) + "</td>\n")
                    outfile.write("</tr>")
            outfile.write("</table>\n\n")
            
            outfile.write("<br>\n<br>--<br>\n")
            outfile.write(self.trUtf8(u"TOTAL DE SEQUÊNCIAS = ") + str(nseq) + self.trUtf8(u" sequência(s)") + "<br>\n\n")
            
            outfile.write("<br>" + self.trUtf8(u"Frequência de CÓDONS") + "<br>\n")
            db_cursor.execute("SELECT " + self.headers[INDIVIDUAL] + ", " +  self.headers[pos] + \
                            " FROM Temp WHERE " + self.headers[pos] + " IS NOT NULL" + \
                            " ORDER BY " + self.headers[INDIVIDUAL])
            
            while True:
                row = db_cursor.fetchone()
                if row == None: break
                seq = str(row[1])
                if len(seq) > 0:
                    outfile.write("\n<br><table border=1 cellspacing=1 cellpadding=1 width=""100%"">\n")
                    outfile.write("<tr><th>" + self.trUtf8(u"INDIVÍDUO") +"</th>")
                    outfile.write("<th>"+ self.trUtf8(u"CÓDON") +"</th>")
                    outfile.write("<th>" + self.trUtf8(u"FREQ.") + "</th>")
                    outfile.write("<th>%</th>")
                    outfile.write("</tr>\n")
                    nuc = DNA(seq)
                    codons = DNA.codons(nuc)
                    dups = DNA.frequency(nuc, codons)
                    cdslen = len(dups)
                    outfile.write("<tr>")	
                    outfile.write("<td align=""Center"">" + str(row[0]) + "</td>\n")
                    outfile.write("<td align=""Left"">&nbsp;</td>\n")
                    outfile.write("<td align=""Left"">&nbsp;</td>\n")
                    outfile.write("<td align=""Left"">&nbsp;</td>\n")
                    outfile.write("</tr>")
                    for i in range(cdslen):
                        outfile.write("<tr>")
                        outfile.write("<td align=""Left"">&nbsp;</td>\n")
                        outfile.write("<td align=""Center"">" + dups[i][0] + "</td>\n")
                        outfile.write("<td align=""Center"">" + str(dups[i][1]) + "</td>\n")
                        outfile.write("<td align=""Center"">" + "{:.1f}".format((dups[i][1] * 100.0 / cdslen)) + "</td>\n")
                        outfile.write("</tr>")
                    outfile.write("</table>\n\n")
                    outfile.write("<br>\n<br>--<br>\n")
                    outfile.write(self.trUtf8(u"TOTAL DE CÓDONS = ") + str(cdslen) + self.trUtf8(u" códon(s)") + "<br>\n")
        
        db_cursor.close()
        db_connection.close()
        
        outfile.write("</body>\n")
        outfile.write("</html>\n")
        outfile.close()
        self.setSource(QtCore.QUrl(self.title))    
    
    def Check(self):
        if not is_online():
            QtGui.QMessageBox.warning(self. self.trUtf8(u"Aviso"), 
                self.trUtf8(u"Não foi detectada uma conexão com a Internet"))
        
        sqlCreateStr = "CREATE TABLE Temp("
        for i in range(len(self.headers)):
            if self.types[i] == "text":
                sqlCreateStr += quote_identifier(self.headers[i], "replace") + " TEXT"
            elif self.types[i] == "numeric":
                sqlCreateStr += quote_identifier(self.headers[i], "replace") + " NUMERIC"
            if i < len(self.headers)-1: sqlCreateStr += ", "
        sqlCreateStr = sqlCreateStr + ")"

        sqlInsertStr = "INSERT INTO Temp VALUES("
        for i in range(len(self.headers)):
            sqlInsertStr += "?"
            if i < len(self.headers)-1: sqlInsertStr += ", "
        sqlInsertStr = sqlInsertStr + ")"
        
        sqlSelectStr = "SELECT DISTINCT TRIM(" + self.headers[SPECIES] + ") FROM Temp" 
        if len(self.filter) > 0: 
            sqlSelectStr += self.filter
        sqlSelectStr += " ORDER BY TRIM(" + self.headers[SPECIES] + ")"
        
        db_data = tuple(self.data)
        db_connection = sqlite3.connect(":memory:")
        db_cursor = db_connection.cursor()
        db_cursor.execute(sqlCreateStr)
        db_cursor.executemany(sqlInsertStr, db_data)
        db_cursor.execute(sqlSelectStr)
        all_rows = db_cursor.fetchall()
        db_cursor.close()
        db_connection.close()
        
        outfile = open(unicode(self.filename), 'w')
        self.Header(outfile, self.trUtf8(u"RELATÓRIO DE VERIFICAÇÃO NOMENCLATURAL"))
        
        v = 0
        s = 0
        x = 0
        c = 0
        q = 0
        
        names = []
        for i in range(len(self.data)):
            names.append(self.data[i][3])
        names.sort()
        names = remove_duplicates(names)
        outfile.write("<br>" + self.trUtf8(u"NOMES SUSPEITOS") + "<br><br>\n")
        for n in names:
            item = process.extract(n, names, limit=2)[1]
            if item[1] > 90:
                outfile.write("<i>" + item[0] + "</i><br>\n")
                x += 1
        outfile.write("\n\n")
        outfile.write("<br>" + self.trUtf8(u"NOMES VERIFICADOS") + "<br>\n")
        
        if is_online():
            progress = QtGui.QProgressDialog(self.trUtf8(u"Verificando nomes..."), 
                            self.trUtf8(u"Cancelar"), 0, len(all_rows))
            progress.setWindowModality(QtCore.Qt.WindowModal)
            for row in all_rows:
                progress.setValue(c)
                if progress.wasCanceled():
                    break
                if row == None: break
                genus, cf, species, author1, subsp, infraname, author2 = parse_name(str(row[0]))
                if genus in [self.trUtf8(u"Indet."), self.trUtf8(u"Desconhecida"), self.trUtf8(u"Morta")] or \
                    species in [self.trUtf8(u"indet."), self.trUtf8(u"desconhecida"), self.trUtf8(u"morta"), "sp.", "spp."]:
                    indet = True
                else:
                    indet = False
                queryStr = genus + " " + species + iif(subsp is not None, " " + infraname, "")
                outfile.write("<br><i>" + queryStr + "</i>")
                if checkCoL(queryStr) and not indet:
                    (name, author, status, valid_name, valid_author, taxon_list) = searchCoL(queryStr)
                    outfile.write(" " + author + " ")
                    if status == "synonym":
                        s += 1
                        outfile.write(self.trUtf8(u" -- sinônimo de ") + "<i>" + valid_name + "</i> " + valid_author)
                    else:
                        v += 1
                        outfile.write(self.trUtf8(u" -- nome válido"))
                    outfile.write("<br>\n&nbsp;&nbsp;&nbsp;")
                    outfile.write("; ".join(taxon_list))
                    outfile.write("\n<br>")
                else:
                    outfile.write(self.trUtf8(u" -- sem registro") + "\n<br>")
                c += 1
                
            q = v / len(all_rows)
        
        outfile.write("\n<br>--<br>\n")
        outfile.write(self.trUtf8(u"TOTAL DE NOMES = ") + str(len(all_rows)) + "<br>\n")
        outfile.write(self.trUtf8(u"TOTAL DE NOMES VÁLIDOS = ") + str(v) + "<br>\n")		
        outfile.write(self.trUtf8(u"TOTAL DE SINÔNIMOS = ") + str(s) + "<br>\n")
        outfile.write(self.trUtf8(u"TOTAL DE NOMES SUSPEITOS = ") + str(x) + "<br>\n")
        outfile.write(self.trUtf8(u"ÍNDICE DE QUALIDADE TAXONÔMICA = ") + "{:.3f}".format(q) + "<br>\n")
        
        outfile.write("</body>\n")
        outfile.write("</html>\n")
        outfile.close()
        self.setSource(QtCore.QUrl(self.title))
        
    def Georef(self):
        if not is_online():
            QtGui.QMessageBox.warning(self. self.trUtf8(u"Aviso"), 
                self.trUtf8(u"Não foi detectada uma conexão com a Internet"))
                
        sqlCreateStr = "CREATE TABLE Temp("
        for i in range(len(self.headers)):
            if self.types[i] == "text":
                sqlCreateStr += quote_identifier(self.headers[i], "replace") + " TEXT"
            elif self.types[i] == "numeric":
                sqlCreateStr += quote_identifier(self.headers[i], "replace") + " NUMERIC"
            if i < len(self.headers)-1: sqlCreateStr += ", "
        sqlCreateStr = sqlCreateStr + ")"

        sqlInsertStr = "INSERT INTO Temp VALUES("
        for i in range(len(self.headers)):
            sqlInsertStr += "?"
            if i < len(self.headers)-1: sqlInsertStr += ", "
        sqlInsertStr = sqlInsertStr + ")"
        
        sqlSelectStr = "SELECT " + self.headers[INDIVIDUAL] + ", TRIM(" + self.headers[SPECIES] + "), TRIM(" + \
                        self.headers[LOCALITY] + "), " + self.headers[LATITUDE] + ", " + \
                        self.headers[LONGITUDE] + " FROM Temp"
        if len(self.filter) > 0: 
            sqlSelectStr += self.filter
        sqlSelectStr += " ORDER BY TRIM(" + self.headers[LOCALITY] + ")"
        
        db_data = tuple(self.data)
        db_connection = sqlite3.connect(":memory:")
        db_cursor = db_connection.cursor()
        db_cursor.execute(sqlCreateStr)
        db_cursor.executemany(sqlInsertStr, db_data)
        db_cursor.execute(sqlSelectStr)
        all_rows = db_cursor.fetchall()
        db_cursor.close()
        db_connection.close()
        
        outfile = open(unicode(self.filename), 'w')
        self.Header(outfile, self.trUtf8(u"RELATÓRIO DE GEOCODIFICAÇÃO"))
        outfile.write("<br><br>\n")
        outfile.write("<table border=0 cellspacing=1 cellpadding=1 width=""100%"">\n")
        outfile.write("<tr><th>" + self.trUtf8(u"INDIVÍDUO") + "</th>")
        outfile.write("<th>"+ self.trUtf8(u"ESPÉCIE") + "</th>")
        outfile.write("<th>" + self.trUtf8(u"LOCALIDADE") + "</th>")
        outfile.write("<th>" + self.trUtf8(u"LATITUDE") + "</th>")
        outfile.write("<th>" + self.trUtf8(u"LONGITUDE") + "</th>")
        outfile.write("<th>" + self.trUtf8(u"ALTITUDE (m)") + "</th>")
        outfile.write("</tr>\n")
        
        coded = 0
        uncoded = 0
        nrows = 0
        
        if is_online():
            progress = QtGui.QProgressDialog(self.trUtf8(u"Verificando coordenadas..."),
                            self.trUtf8(u"Cancelar"), 0, len(all_rows))
            progress.setWindowModality(QtCore.Qt.WindowModal)
            for row in all_rows:
                progress.setValue(nrows)
                if progress.wasCanceled():
                    break
                if row == None: break
                Country = self.projectdata["country"]
                State = self.projectdata["state"]
                CollectionCode = str(row[0])
                Species = str(row[1])
                Locality = str(row[2])
                Latitude = to_float(row[3])
                Longitude = to_float(row[4])
                if abs(Latitude) == 0.0 and abs(Longitude) == 0.0:
                    georef = False
                    uncoded += 1
                else:
                    georef = True
                    coded += 1
                if not georef and len(Locality) > 0:
                    genus, cf, species, author1, subsp, infraname, author2 = parse_name(Species)
                    geolocator = GoogleV3()
                    if len(State) > 0:
                        loc = Locality + ',' + State + ',' + Country
                    else:
                        loc = Locality + ',' + Country
                    try:
                        location = geolocator.geocode(loc)
                        outfile.write("<tr>")	
                        outfile.write("<td align=""Left"">" + CollectionCode + "</td>\n")
                        outfile.write("<td align=""Left""><i>" + genus + ' ' + species + "</i></td>\n")
                        outfile.write("<td align=""Left"">" + unicode_to_ascii(location.address) + "</td>\n")
                        outfile.write("<td align=""Center"">" + str(location.latitude) + "</td>\n")
                        outfile.write("<td align=""Center"">" + str(location.longitude) + "</td>\n")
                        outfile.write("<td align=""Center"">" + str(location.altitude) + "</td>\n")
                        outfile.write("</tr>")
                    except geopy.exc.GeocoderTimedOut:
                        break
                nrows += 1
        
        outfile.write("</table>\n\n")
        outfile.write("\n<br>--<br>\n")
        outfile.write(self.trUtf8(u"TOTAL DE REGISTROS = ") + str(len(all_rows)) + "<br>\n")
        outfile.write(self.trUtf8(u"TOTAL DE REGISTROS GEOCODIFICADOS = ") + str(coded) + "<br>\n")		
        outfile.write(self.trUtf8(u"TOTAL DE REGISTROS NÃO-GEOCODIFICADOS = ") + str(uncoded) + "<br>\n")
        
        outfile.write("</body>\n")
        outfile.write("</html>\n")
        outfile.close()
        self.setSource(QtCore.QUrl(self.title))
        
    def Divers(self, datmat, graph_it):
        sqlCreateStr = "CREATE TABLE Temp("
        for i in range(len(self.headers)):
            if self.types[i] == "text":
                sqlCreateStr += quote_identifier(self.headers[i], "replace") + " TEXT"
            elif self.types[i] == "numeric":
                sqlCreateStr += quote_identifier(self.headers[i], "replace") + " NUMERIC"
            if i < len(self.headers)-1: sqlCreateStr += ", "
        sqlCreateStr = sqlCreateStr + ")"

        sqlInsertStr = "INSERT INTO Temp VALUES("
        for i in range(len(self.headers)):
            sqlInsertStr += "?"
            if i < len(self.headers)-1: sqlInsertStr += ", "
        sqlInsertStr = sqlInsertStr + ")"
        
        db_data = tuple(self.data)
        db_connection = sqlite3.connect(":memory:")
        db_cursor = db_connection.cursor()
        db_cursor.execute(sqlCreateStr)
        db_cursor.executemany(sqlInsertStr, db_data)
        
        outfile = open(unicode(self.filename), 'w')
        self.Header(outfile, self.trUtf8(u"ANÁLISE DE DIVERSIDADE"))
        outfile.write("<br>\n")
        
        outfile.write("<table border=1 cellspacing=1 cellpadding=1 width=""100%"">\n")
        outfile.write("<tr><th>" + self.trUtf8(u"LOCAL") + "</th>")
        outfile.write("<th>N0 (S)</th>")
        outfile.write("<th>H'</th>")
        outfile.write("<th>N1</th>")
        outfile.write("<th>1-&lambda;</th>")
        outfile.write("<th>N2 (1/&lambda;)</th>")
        outfile.write("</tr>\n")
            
        db_cursor.execute("SELECT DISTINCT TRIM(" + self.headers[SAMPLE] + ") FROM Temp")
        l = len(db_cursor.fetchall())
            
        db_cursor.execute("SELECT DISTINCT TRIM(" + self.headers[SPECIES] + ") FROM Temp")
        s = len(db_cursor.fetchall())
            
        db_cursor.execute("SELECT DISTINCT TRIM(" + self.headers[SAMPLE] + ") FROM Temp ORDER BY TRIM(" + 
                self.headers[SAMPLE] + ")")
                
        if graph_it:
            spp = []
                
        D, indices = sample_diversity(datmat, indices=["shannon", "hill", "simpson", "simpson_inv"])
        
        i = 0
        while True:
            row = db_cursor.fetchone()
            if row == None: break
            vloc = str(row[0])
            aux_cursor = db_connection.cursor()
            aux_cursor.execute("SELECT COUNT(DISTINCT TRIM(" + self.headers[SPECIES] + ")) FROM Temp WHERE TRIM(" + \
                        self.headers[SAMPLE] + ") = '" + vloc + "'")
            aux_row = aux_cursor.fetchone()
            nspp = int(aux_row[0])
            outfile.write("<tr>")	
            outfile.write("<td align=""Left"">" + vloc + "</td>\n")
            outfile.write("<td align=""Center"">" + str(nspp) + "</td>\n")
            outfile.write("<td align=""Center"">" + "{:.3f}".format(D[i,0]) + "</td>\n")
            outfile.write("<td align=""Center"">" + "{:.3f}".format(D[i,1]) + "</td>\n")
            outfile.write("<td align=""Center"">" + "{:.3f}".format(D[i,2]) + "</td>\n")
            outfile.write("<td align=""Center"">" + "{:.1f}".format(D[i,3]) + "</td>\n")
            outfile.write("</tr>")
            aux_cursor.close()
            if graph_it:
                spp.append(nspp)
            i += 1
        outfile.write("</table>\n\n")	
        outfile.write("<br>\n<br>--<br>\n")
        outfile.write(self.trUtf8(u"TOTAL DE LOCAIS = ") + str(l) + self.trUtf8(u" locais") + "<br>\n")
        outfile.write(self.trUtf8(u"TOTAL DE ESPÉCIES = ") + str(s) + self.trUtf8(u" espécie(s)") + "<br>\n")
            
        if graph_it:
            freq = numpy.array(spp)
            ES, S, N = rarefact(freq)
            n = numpy.arange(N)
            figf = os.path.splitext(os.path.basename(str(self.filename)))[0] + ".png"
            plt.clf()
            plt.plot(n, ES, color="red", linewidth=3.0, alpha=0.6)
            plt.title(self.trUtf8(u"CURVA DE RAREFAÇÃO"))
            plt.xlabel(self.trUtf8(u"TAMANHO DA AMOSTRA"))
            plt.ylabel(self.trUtf8(u"NÚMERO ESPERADO ") + "($E[S_{n}]$)")
            plt.grid(False)
            plt.tight_layout()
            plt.savefig(figf, dpi=72)
            outfile.write("<img src='" + figf + "'>\n")
        
        db_cursor.close()
        db_connection.close()
        
        outfile.write("</body>\n")
        outfile.write("</html>\n")
        outfile.close()
        self.setSource(QtCore.QUrl(self.title))
        
    def Cluster(self, datmat, transf, coef, method, graph_it):
        simil = [self.trUtf8(u"Distância de Bray-Curtis"), 
                self.trUtf8(u"Métrica de Canberra"), 
                self.trUtf8(u"Distância Manhattan"),
                self.trUtf8(u"Distância Euclidiana simples"),
                self.trUtf8(u"Distância Euclidiana normalizada"),
                self.trUtf8(u"Distância Euclidiana quadrada"),
                self.trUtf8(u"Distância de Morisita-Horn"),
                self.trUtf8(u"Correlação linear de Pearson"),
                self.trUtf8(u"Similaridade de Jaccard"),
                self.trUtf8(u"Similaridade de Dice-Sorenson"),
                self.trUtf8(u"Similaridade de Kulczynski"),
                self.trUtf8(u"Similaridade de Ochiai")]
        methods = ["SLM", "CLM", "UPGMA", "WPGMA", "UPGMC", "WPGMC", "Ward"]
        nobs, nvar = datmat.shape
        outfile = open(unicode(self.filename), 'w')
        self.Header(outfile, self.trUtf8(u"ANÁLISE DE AGRUPAMENTOS"))
        outfile.write("<br>\n")
        outfile.write(str(nobs) + self.trUtf8(u" amostras") + " x " + str(nvar) + self.trUtf8(u" espécies") + "<br><br>\n")
        
        samples = []
        for i in range(len(self.data)):
            samples.append(self.data[i][0])
        samples = list(set(samples))
        samples = sorted(samples)
        
        if transf == 1:
            datamat = numpy.log10(datmat)
            outfile.write(self.trUtf8(u"Transformação: Logaritmo comum (base 10)") + "<br><br>\n")
        elif transf == 2:
            datmat == numpy.log(datmat)
            outfile.write(self.trUtf8(u"Transformação: Logaritmo natural (base e)") + "<br><br>\n")
        elif transf == 3:
            datmat = numpy.sqrt(datmat)
            outfile.write(self.trUtf8(u"Transformação: Raiz quadrada") + "<br><br>\n")
        elif transf == 4: 
            datmat = numpy.arcsin(datmat)
            outfile.write(self.trUtf8(u"Transformação: Arcosseno") + "<br><br>\n")
        else:
            outfile.write(self.trUtf8(u"Sem Transformação") + "<br><br>\n")
        
        if coef == 0:
            dist = pdist(datmat, "braycurtis")
            outfile.write(self.trUtf8(u"Coeficiente: ") + simil[coef] + "<br><br>\n")
        elif coef == 1:
            dist = pdist(datmat, "canberra")
            outfile.write(self.trUtf8(u"Coeficiente: ") + simil[coef] + "<br><br>\n")
        elif coef == 2:
            dist = pdist(datmat, "cityblock")
            outfile.write(self.trUtf8(u"Coeficiente: ") + simil[coef] + "<br><br>\n")
        elif coef == 3:
            dist = pdist(datmat, "euclidean")
            outfile.write(self.trUtf8(u"Coeficiente: ") + simil[coef] + "<br><br>\n")
        elif coef == 4:
            dist = pdist(datmat, "seuclidean")
            outfile.write(self.trUtf8(u"Coeficiente: ") + simil[coef] + "<br><br>\n")
        elif coef == 5:
            dist = pdist(datmat, "sqeuclidean")
            outfile.write(self.trUtf8(u"Coeficiente: ") + simil[coef] + "<br><br>\n")
        elif coef == 6:
            dist = squareform(morisita_horn(datmat))
            outfile.write(self.trUtf8(u"Coeficiente: ") + simil[coef] + "<br><br>\n")
        elif coef == 7:
            dist = pdist(datmat, "correlation")
            outfile.write(self.trUtf8(u"Coeficiente: ") + simil[coef] + "<br><br>\n")
        elif coef == 8:
            dist = pdist(datmat, "jaccard")
            outfile.write(self.trUtf8(u"Coeficiente: ") + simil[coef] + "<br><br>\n")
        elif coef == 9:
            dist = pdist(datmat, "dice")
            outfile.write(self.trUtf8(u"Coeficiente: ") + simil[coef] + "<br><br>\n")
        elif coef == 10:
            dist = pdist(datmat, "kulsinski")
            outfile.write(self.trUtf8(u"Coeficiente: ") + simil[coef] + "<br><br>\n")
        elif coef == 11:
            dist = squareform(ochiai(datmat))
            outfile.write(self.trUtf8(u"Coeficiente: ") + simil[coef] + "<br><br>\n")
        
        if method == 0:
            clus = linkage(dist, "single")
            outfile.write(self.trUtf8(u"Método: Ligação Simples (SLM)") + "<br><br>\n")
        elif method == 1:
            clus = linkage(dist, "complete")
            outfile.write(self.trUtf8(u"Método: Ligação Completa (CLM)") + "<br><br>\n")
        elif method == 2:
            clus = linkage(dist, "average")
            outfile.write(self.trUtf8(u"Método: Média Não-Ponderada (UPGMA)") + "<br><br>\n")
        elif method == 3:
            clus = linkage(dist, "weighted")
            outfile.write(self.trUtf8(u"Método: Média Ponderada (WPGMA)") + "<br><br>\n")
        elif method == 4:
            clus = linkage(datmat, "centroid")
            outfile.write(self.trUtf8(u"Método: Centroide (UPGMC)") + "<br><br>\n")
        elif method == 5:
            clus = linkage(datmat, "median")
            outfile.write(self.trUtf8(u"Método: Mediana (WPGMC)") + "<br><br>\n")
        elif method == 6:
            clus = linkage(datmat, "ward")
            outfile.write(self.trUtf8(u"Método: Variância Mínima de Ward") + "<br><br>\n")
        
        corr = cophenet(clus, dist)[0]
        outfile.write(self.trUtf8(u"Correlação cofenética = ") + "{:.4f}".format(corr) + "<br>\n")
        
        if graph_it:
            figf = os.path.splitext(os.path.basename(str(self.filename)))[0] + ".png"
            plt.clf()
            dendrogram(clus, labels=samples, orientation="right")
            plt.title(self.trUtf8(u"Dendrograma ") + '(' + methods[method] + ')')
            plt.ylabel(self.trUtf8(u"Amostras"))
            plt.xlabel(simil[coef])
            plt.tight_layout()
            plt.savefig(figf, dpi=72)
            outfile.write("<img src='" + figf + "'>\n")
        
        outfile.write("</body>\n")
        outfile.write("</html>\n")
        outfile.close()
        self.setSource(QtCore.QUrl(self.title))
        
    def Ord(self, datmat, envmat, transf, method, coef, index, center, scale, iter, config, scaling, mask, graph_it):
        methods = [self.trUtf8(u"ANÁLISE DE COMPONENTES PRINCIPAIS"),
                    self.trUtf8(u"ANÁLISE DE COORDENADAS PRINCIPAIS"),
                    self.trUtf8(u"ESCALONAMENTO MULTIDIMENSIONAL NÃO-MÉTRICO"),
                    self.trUtf8(u"ANÁLISE DE CORRESPONDÊNCIAS"),
                    self.trUtf8(u"ANÁLISE DE REDUNDÂNCIAS"),
                    self.trUtf8(u"ANÁLISE DE CORRESPONDÊNCIAS CANÔNICA")]
        simil = [self.trUtf8(u"Distância de Bray-Curtis"), 
                self.trUtf8(u"Métrica de Canberra"), 
                self.trUtf8(u"Distância Manhattan"),
                self.trUtf8(u"Distância Euclidiana simples"),
                self.trUtf8(u"Distância Euclidiana normalizada"),
                self.trUtf8(u"Distância Euclidiana quadrada"),
                self.trUtf8(u"Distância de Morisita-Horn")]
        nobs, nvar = datmat.shape
        outfile = open(unicode(self.filename), 'w')
        self.Header(outfile, methods[method])
        outfile.write("<br>\n")
        outfile.write(str(nobs) + self.trUtf8(u" amostras") + " x " + str(nvar) + self.trUtf8(u" espécies") + "<br><br>\n")
        
        samples = []
        for i in range(len(self.data)):
            samples.append(self.data[i][0])
        samples = list(set(samples))
        samples = sorted(samples)
        
        species = []
        for j in range(len(self.data)):
            species.append(str(j+1))
        species = sorted(species)
        
        if transf == 1:
            datamat = numpy.log10(datmat)
            outfile.write(self.trUtf8(u"Transformação: Logaritmo comum (base 10)") + "<br><br>\n")
        elif transf == 2:
            datmat == numpy.log(datmat)
            outfile.write(self.trUtf8(u"Transformação: Logaritmo natural (base e)") + "<br><br>\n")
        elif transf == 3:
            datmat = numpy.sqrt(datmat)
            outfile.write(self.trUtf8(u"Transformação: Raiz quadrada") + "<br><br>\n")
        elif transf == 4: 
            datmat = numpy.arcsin(datmat)
            outfile.write(self.trUtf8(u"Transformação: Arcosseno") + "<br><br>\n")
        else:
            outfile.write(self.trUtf8(u"Sem Transformação") + "<br><br>\n")
            
        if coef == 0:
            dist = squareform(pdist(datmat, "braycurtis"))
        elif coef == 1:
            dist = squareform(pdist(datmat, "canberra"))
        elif coef == 2:
            dist = squareform(pdist(datmat, "cityblock"))
        elif coef == 3:
            dist = squareform(pdist(datmat, "euclidean"))
        elif coef == 4:
            dist = squareform(pdist(datmat, "seuclidean"))
        elif coef == 5:
            dist = squareform(pdist(datmat, "sqeuclidean"))
        elif coef == 6:
            dist = squareform(morisita_horn(datmat))
        
        if method == 0:
            eig_val, eig_vec, scores, sumvariance, cumvariance = pca(datmat, index=index, \
                center=center, stand=scale)
            if center:
                outfile.write(self.trUtf8(u"Dados Centrados") + "<br><br>\n")
            if scale:
                outfile.write(self.trUtf8(u"Dados Estandardizados") + "<br><br>\n")
        elif method == 1:
            eig_val, eig_vec, scores, sumvariance, cumvariance = pcoa(dist)
            outfile.write(self.trUtf8(u"Coeficiente: ") + simil[coef] + "<br><br>\n")
        elif method == 2:
            nm = nmds(dist, initial_pts=config, max_iterations=iter)
            pts = nm.getPoints()
            stress = nm.getStress()
            dim = nm.getDimension()
            outfile.write(self.trUtf8(u"Coeficiente: ") + simil[coef] + "<br><br>\n")
        elif method == 3:
            eig_val, row_scores, col_scores, sumvariance, cumvariance = ca(datmat, scaling)
        elif method == 4:
            if len(mask) > 0:
                newmat = envmat[:,[mask]]
                newmat = numpy.reshape(newmat,(nobs, len(mask)))
            else:
                newmat = envmat
            eig_val, row_scores, col_scores, biplot_scores, sumvariance, cumvariance, \
                site_constraints = rda(datmat, newmat, True, 2)
        elif method == 5:
            if len(mask) > 0:
                newmat = envmat[:,[mask]]
                newmat = numpy.reshape(newmat,(nobs, len(mask)))
            else:
                newmat = envmat
            eig_val, row_scores, col_scores, biplot_scores, sumvariance, cumvariance, \
                site_constraints = cca(datmat, newmat, scaling)
                        
        if method == 0 or method == 1:
            nvect = 0
            outfile.write(self.trUtf8(u"AUTOVALORES") + "<br>\n")
            outfile.write("<table border=1 cellspacing=1 cellpadding=1 width=""100%"">\n")
            outfile.write("<tr><th>i</th>")
            outfile.write("<th>" + self.trUtf8(u"AUTOVALOR") + "</th>")
            outfile.write("<th>" + self.trUtf8(u"%VARIÂNCIA") + "</th>")
            outfile.write("<th>" + self.trUtf8(u"%CUMULATIVA") + "</th>")
            outfile.write("</tr>\n")
            for k in range(0, nvar):
                if eig_val[k] < 0.0001: break
                outfile.write("<tr><td align=""Center"">" + str(k + 1) + "</td>\n")
                outfile.write("<td align=""Center"">" + "{:.3f}".format(eig_val[k]) + "</td>\n")
                outfile.write("<td align=""Center"">" + "{:.3f}".format(sumvariance[k]) + "</td>\n")
                outfile.write("<td align=""Center"">" + "{:.3f}".format(cumvariance[k]) + "</td>\n")
                outfile.write("</tr>\n")
                nvect += 1
            outfile.write("</table><br><br>\n\n")
        
            outfile.write(self.trUtf8(u"AUTOVETORES") + "<br>\n")
            outfile.write("<table border=1 cellspacing=1 cellpadding=1 width=""100%"">\n")
            outfile.write("<tr><th>" + self.trUtf8(u"VARIÁVEL") + "</th>")
            for i in range(nvect):
                outfile.write("<th>" + self.trUtf8(u"EIXO ") + str(i + 1) + "</th>")
            outfile.write("</tr>\n")
            for i in range(len(eig_vec)):
                outfile.write("<tr>")
                outfile.write("<td align=""Center"">" + str(i + 1) + "</td>\n")
                for j in range(0, nvect):
                    outfile.write("<td align=""Center"">" + "{:.3f}".format(eig_vec[i][j]) + "</td>\n")
                outfile.write("</tr>\n")
            outfile.write("</table>\n\n")
            
        elif method == 2:
            outfile.write(self.trUtf8(u"Stress = ") + "{:.8f}".format(stress) + "<br><br>\n")
            outfile.write(self.trUtf8(u"COORDENADAS") + "<br>\n")
            outfile.write("<table border=1 cellspacing=1 cellpadding=1 width=""100%"">\n")
            outfile.write("<tr><th>" + self.trUtf8(u"AMOSTRA") + "</th>")
            for i in range(dim):
                outfile.write("<th>" + self.trUtf8(u"DIMENSÃO ") + str(i + 1) + "</th>")
            outfile.write("</tr>\n")
            for i in range(len(pts)):
                outfile.write("<tr>")
                outfile.write("<td align=""Center"">" + str(i + 1) + "</td>\n")
                for j in range(0, dim):
                    outfile.write("<td align=""Center"">" + "{:.3f}".format(pts[i][j]) + "</td>\n")
                outfile.write("</tr>\n")
            outfile.write("</table>\n\n")
        
        elif method == 3 or method == 4 or method == 5:
            nvect = 0
            outfile.write(self.trUtf8(u"AUTOVALORES") + "<br>\n")
            outfile.write("<table border=1 cellspacing=1 cellpadding=1 width=""100%"">\n")
            outfile.write("<tr><th>i</th>")
            outfile.write("<th>" + self.trUtf8(u"AUTOVALOR") + "</th>")
            outfile.write("<th>" + self.trUtf8(u"%VARIÂNCIA") + "</th>")
            outfile.write("<th>" + self.trUtf8(u"%CUMULATIVA") + "</th>")
            outfile.write("</tr>\n")
            for k in range(0, len(eig_val)):
                if eig_val[k] < 0.0001: break
                outfile.write("<tr><td align=""Center"">" + str(k + 1) + "</td>\n")
                outfile.write("<td align=""Center"">" + "{:.3f}".format(eig_val[k]) + "</td>\n")
                outfile.write("<td align=""Center"">" + "{:.3f}".format(sumvariance[k]) + "</td>\n")
                outfile.write("<td align=""Center"">" + "{:.3f}".format(cumvariance[k]) + "</td>\n")
                outfile.write("</tr>\n")
                nvect += 1
            outfile.write("</table><br><br>\n\n")
            
            outfile.write(self.trUtf8(u"ESCORES DAS AMOSTRAS") + "<br>\n")
            outfile.write("<table border=1 cellspacing=1 cellpadding=1 width=""100%"">\n")
            outfile.write("<tr><th>" + self.trUtf8(u"AMOSTRA") + "</th>")
            for i in range(nvect):
                outfile.write("<th>" + self.trUtf8(u"EIXO ") + str(i + 1) + "</th>")
            outfile.write("</tr>\n")
            for i in range(len(row_scores)):
                outfile.write("<tr>")
                outfile.write("<td align=""Center"">" + str(i + 1) + "</td>\n")
                for j in range(nvect):
                    outfile.write("<td align=""Center"">" + "{:.3f}".format(row_scores[i][j]) + "</td>\n")
                outfile.write("</tr>\n")
            outfile.write("</table><br><br>\n\n")
            
            outfile.write(self.trUtf8(u"ESCORES DAS VARIÁVEIS") + "<br>\n")
            outfile.write("<table border=1 cellspacing=1 cellpadding=1 width=""100%"">\n")
            outfile.write("<tr><th>" + self.trUtf8(u"VARIÁVEL") + "</th>")
            for i in range(nvect):
                outfile.write("<th>" + self.trUtf8(u"EIXO ") + str(i + 1) + "</th>")
            outfile.write("</tr>\n")
            for i in range(len(col_scores)):
                outfile.write("<tr>")
                outfile.write("<td align=""Center"">" + str(i + 1) + "</td>\n")
                for j in range(nvect):
                    outfile.write("<td align=""Center"">" + "{:.3f}".format(col_scores[i][j]) + "</td>\n")
                outfile.write("</tr>\n")
            outfile.write("</table><br><br>\n\n")
        
        if graph_it:
            if method == 0 or method == 1:
                figf = os.path.splitext(os.path.basename(str(self.filename)))[0] + ".png"
                plt.clf()
                plt.scatter(scores[:,0], scores[:,1])
                for label, x, y in zip(samples, scores[:,0], scores[:,1]):
                    plt.annotate(label, xy = (x, y), xytext = (-10, 10), textcoords="offset points")
                plt.title(self.trUtf8(u"DIAGRAMA DE DISPERSÃO"))
                plt.xlabel(self.trUtf8(u"EIXO 1") + " (" + "{:.1f}".format(sumvariance[0]) + " %)")
                plt.ylabel(self.trUtf8(u"EIXO 2") + " (" + "{:.1f}".format(sumvariance[1]) + " %)")
                plt.tight_layout()
                plt.savefig(figf, dpi=72)
                outfile.write("<img src='" + figf + "'>\n")
            
            elif method == 2:
                figf = os.path.splitext(os.path.basename(str(self.filename)))[0] + ".png"
                plt.clf()
                plt.scatter(pts[:,0], pts[:,1])
                for label, x, y in zip(samples, pts[:,0], pts[:,1]):
                    plt.annotate(label, xy = (x, y), xytext = (-10, 10), textcoords="offset points")
                plt.title(self.trUtf8(u"DIAGRAMA DE DISPERSÃO"))
                plt.xlabel(self.trUtf8(u"DIMENSÃO 1"))
                plt.ylabel(self.trUtf8(u"DIMENSÃO 2"))
                plt.tight_layout()
                plt.savefig(figf, dpi=72)
                outfile.write("<img src='" + figf + "'>\n")
            
            elif method == 3 or method == 4 or method == 5:
                figf = os.path.splitext(os.path.basename(str(self.filename)))[0] + ".png"
                plt.clf()
                plt.scatter(row_scores[:,0], row_scores[:,1], marker='o', color="blue")
                plt.scatter(col_scores[:,0], col_scores[:,1], marker='^', color="red")
                for label, x, y in zip(samples, row_scores[:,0], row_scores[:,1]):
                    plt.annotate(label, xy = (x, y), xytext = (-5, 5), textcoords="offset points")
                for label, x, y in zip(species, col_scores[:,0], col_scores[:,1]):
                    plt.annotate(label, xy = (x, y), xytext = (-5, 5), textcoords="offset points")
                plt.title(self.trUtf8(u"DIAGRAMA DE DISPERSÃO"))
                plt.xlabel(self.trUtf8(u"EIXO 1") + " (" + "{:.1f}".format(sumvariance[0]) + " %)")
                plt.ylabel(self.trUtf8(u"EIXO 2") + " (" + "{:.1f}".format(sumvariance[1]) + " %)")
                plt.tight_layout()
                plt.savefig(figf, dpi=72)
                outfile.write("<img src='" + figf + "'>\n")
            
        outfile.write("</body>\n")
        outfile.write("</html>\n")
        outfile.close()
        self.setSource(QtCore.QUrl(self.title))
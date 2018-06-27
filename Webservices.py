# -*- coding: utf-8 -*-
#================================================================================#
#  ECOLOG - Sistema Gerenciador de Banco de Dados para Levantamentos Ecológicos  #
#        ECOLOG - Database Management System for Ecological Surveys              #
#      Copyright (c) 1990-2014 Mauro J. Cavalcanti. All rights reserved.         #
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
#    SOAPPy 0.12+ (pywebsvcs.sourceforge.net)                                    #
#================================================================================#

import urllib
from urllib2 import urlopen
from xml.dom import minidom
from SOAPpy import WSDL

from Useful import is_online

def checkCoL(searchStr):
	usock = urllib.urlopen("http://www.catalogueoflife.org/col/webservice?name=" + urllib.quote_plus(searchStr))
	xmldoc = minidom.parse(usock)
	usock.close()
	results = xmldoc.getElementsByTagName("results")
	attribute_list = results[0]
	errmsg = attribute_list.attributes["error_message"].value
	if len(errmsg) == 0:
		return True
	else:
		return False

def searchCoL(searchStr):
	usock = urllib.urlopen("http://www.catalogueoflife.org/col/webservice?name=" + urllib.quote_plus(searchStr) + "&response=full")
	xmldoc = minidom.parse(usock)
	usock.close()
	try:
		#--- get name and status
		name = xmldoc.getElementsByTagName("name")[0].firstChild.data
		author = xmldoc.getElementsByTagName("author")[0].firstChild.data
		status = xmldoc.getElementsByTagName("name_status")[0].firstChild.data
		#--- if name is a synonym, get the accepted name
		if status == "synonym":
			item_node = xmldoc.getElementsByTagName("accepted_name")[0]
			valid_name = item_node.getElementsByTagName("name")[0].firstChild.data
			valid_author = item_node.getElementsByTagName("author")[0].firstChild.data
		else:
			valid_name = name
			valid_author = author
		#--- get higher taxa for this name
		taxon_list = []
		for i in range(5):
			item_node = xmldoc.getElementsByTagName("taxon")[i]
			item = item_node.getElementsByTagName("name")[0].firstChild.data
			taxon_list.append(item)
	except:
			name = ""
			author = ""
			status = ""
			valid_name = ""
			valid_author = ""
			taxon = []
	return (name, author, status, valid_name, valid_author, taxon_list)

def searchWoRMS(searchStr):
	wsdlFile = "http://www.marinespecies.org/aphia.php?p=soap&wsdl=1"
	server = WSDL.Proxy(wsdlFile)
	ID = server.getAphiaID(searchStr, False)
	record = server.getAphiaRecordByID(ID)
	return record
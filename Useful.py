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
#================================================================================#

from __future__ import division
from urllib2 import urlopen
import os, time, sys, math, locale, codecs, unicodedata
from htmlentitydefs import codepoint2name

def alfa(str,ch = '-'):
	return str.replace(ch, '').isalpha() 

def capfirst(value):
	s = (value[0]).upper()
	for i in range(1,len(value)):
		if ord(value[i - 1]) < 33:
			s = s + value[i].upper()
		else:
			s = s + value[i].lower()
	return s

def corr(x, y=None):
	if y is not None:
		if y.shape[0] != x.shape[0]:
			raise ValueError("Both matrices must have the same number of rows")
		x, y = scale(x), scale(y)
	else:
		x = scale(x)
		y = x
	return x.T.dot(y) / x.shape[0]

def degtodms(degfloat):
	deg = int(degfloat)
	minfloat = 60 * (degfloat - deg)
	min = int(minfloat)
	secfloat = 60 * (minfloat - min)
	secfloat = round(secfloat, 3)
	if secfloat == 60:
		min = min + 1
		secfloat = 0
	if min == 60:
		deg = deg + 1
		min = 0
	return deg, abs(min), abs(secfloat)

def extenso(data, mesext):
	if data == "None": return ""
	dia, mes, ano = data.split('-')
	dia = int(dia)
	mes = int(mes)
	return str(dia) + " de " + mesext[mes] + " de " + ano

def extract(text, sub1, sub2):
	return text.split(sub1)[-1].split(sub2)[0]

def find(l, s):
	for i in range(len(l)):
		if l[i].find(s) != -1:
			return i
	return None # Or -1

def force_decode(string, codecs=["utf8", "cp1252"]):
	for i in codecs:
		try:
			return string.decode(i)
		except:
			pass

def get_unit(s):
	if s.find('(') != -1 and s.find(')') != -1:
		return s[s.find("(")+1:s.find(")")]
	else:
		return ""

def htmlescape(text):
	text = (text).decode("utf-8")
	d = dict((unichr(code), u'&%s;' % name) for code,name in codepoint2name.iteritems() if code!=38)    
	if u"&" in text:
		text = text.replace(u"&", u"&amp;")
	for key, value in d.iteritems():
		if key in text:
			text = text.replace(key, value)
	return text

def iif(boolVar, ifTrue, ifFalse):
	if boolVar:
		return ifTrue
	else:
		return ifFalse

def is_digit(s):
	n = abs(eval(s))
	return str(n).isdigit()

def is_number(s):
	try:
		float(s)
		return True
	except ValueError:
		pass
	try:
		import unicodedata
		unicodedata.numeric(s)
		return True
	except (TypeError, ValueError):
		pass
	return False

def is_online(reliableserver="http://www.google.com"):
	try:
		urlopen(reliableserver)
		return True
	except IOError:
		return False

def parse_name(taxonomic_name):
	if taxonomic_name == "" or taxonomic_name == None:
		return ("", "", "", "", "", "", "")
	genus = ""
	aff = ""
	species = ""
	author1 = ""
	ssp = ""
	subspecies = ""
	author2 = ""
	taxonomy = taxonomic_name.split(" ")
	genus = taxonomy.pop(0)
	if len(taxonomy) > 0:
		if taxonomy[0].startswith("cf.") or taxonomy[0].startswith("aff."):
			aff = taxonomy.pop(0)
	if len(taxonomy) > 0:
		species = taxonomy.pop(0)
	if len(taxonomy) > 0:
		author1 = taxonomy.pop(0)
	if len(taxonomy) > 0:
		if taxonomy[0].startswith("var.") or taxonomy[0].startswith("subsp."):
			ssp = taxonomy.pop(0)
	if len(taxonomy) > 0:
		subspecies = taxonomy.pop(0)
	if len(taxonomy) > 0:
		author2 = taxonomy.pop(0)
	return (genus, aff, species, author1, ssp, subspecies, author2)

def percent(val1, val2):
	if val2 == 0:
		ret_val = 0.0
	else:
		ret_val = (float(val1) / float(val2)) * 100.00
	return ret_val

def quote_identifier(s, errors="strict"):
	encodable = s.encode("utf-8", errors).decode("utf-8")
	nul_index = encodable.find("\x00")
	if nul_index >= 0:
		error = UnicodeEncodeError("NUL-terminated utf-8", encodable,
								nul_index, nul_index + 1, "NUL not allowed")
		error_handler = codecs.lookup_error(errors)
		replacement, _ = error_handler(error)
		encodable = encodable.replace("\x00", replacement)
	return "\"" + encodable.replace("\"", "\"\"") + "\""

def remove_duplicates(values):
	output = []
	seen = set()
	for value in values:
		if value not in seen:
			output.append(value)
			seen.add(value)
	return output

def roman(data):
	if data == "None": return ""
	mesext = ["I", "II", "III", "IV", "V", "VI", "VII", "VIII", "IX", "X", "XI", "XII"]
	dia, mes, ano = data.split('-')
	dia = int(dia)
	mes = int(mes)
	return str(dia) + "-" + mesext[mes] + "-" + ano

def strip_letters(str):
	str = str.lower().replace(' ', '')
	for char in str:
		if char.isalpha():
			str = str.replace(char, '')
	return str

def substr(str, source, size):
	return str[source:size+source]

def to_int(in_val):
	try:
		ret_val = int(in_val)
		return ret_val
	except:
		return 0

def to_float(in_val):
	try:
		ret_val = float(in_val)
		return ret_val
	except:
		return 0.0	

def truncate(s, length = 1, etc = "..."):
	if len(s) < length:
		return s
	else:
		return s[:length] + etc
	
def unicode_to_ascii(str):
	return unicodedata.normalize("NFKD", unicode(str)).encode("ascii","ignore")
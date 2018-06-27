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
#    dbfpy 2.0+ (dbfpy.sourceforge.net/)                                         #
#    shapefile 1.2+ (github.com/mlacayoemery/shapefile)                          #
#================================================================================#

import os, time, datetime, sqlite3
from datetime import datetime
from dbfpy import dbf
import shapefile

from Useful import (degtodms, get_unit, iif, parse_name, quote_identifier, strip_letters, to_float, to_int, 
					unicode_to_ascii)

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

def toEML(filename, projectdata, headers, types, data):
	sqlCreateStr = "CREATE TABLE Temp("
	for i in range(len(headers)):
		if types[i] == "text":
			sqlCreateStr += quote_identifier(headers[i], "replace") + " TEXT"
		elif types[i] == "numeric":
			sqlCreateStr += quote_identifier(headers[i], "replace") + " NUMERIC"
		if i < len(headers) - 1: sqlCreateStr += ", "
	sqlCreateStr += ")"

	sqlInsertStr = "INSERT INTO Temp VALUES("
	for i in range(len(headers)):
		sqlInsertStr += "?"
		if i < len(headers) - 1: sqlInsertStr += ", "
	sqlInsertStr += ")"
		
	db_data = tuple(data)
	db_connection = sqlite3.connect(":memory:")
	db_cursor = db_connection.cursor()
	db_cursor.execute(sqlCreateStr)
	db_cursor.executemany(sqlInsertStr, db_data)
	now = datetime.now()
	xmldoc = ""
	levels = 0
	
	emlHeader = "<?xml version=\"1.0\" encoding=\"UTF-8\"?>\n<eml:eml\n \
		packageId=\"eml.1.1\" system=\"knb\"\n \
		xmlns:eml=\"eml://ecoinformatics.org/eml-2.1.0\"\n \
		xmlns:xsi=\"http://www.w3.org/2001/XMLSchema-instance\"\n \
		xmlns:stmml=\"http://www.xml-cml.org/schema/stmml-1.1\"\n \
		xsi:schemaLocation=\"eml://ecoinformatics.org/eml-2.1.0 eml.xsd\">\n \
	<dataset>\n"
	xmldoc += emlHeader

	titulo = projectdata["title"]
	descricao = projectdata["description"]
	pais = projectdata["country"]
	estado = projectdata["state"]
	municipio = projectdata["province"]
	localidade = projectdata["locality"]
	#keywords = str(row["KEYWORDS"]).split(',')
	responsavel = projectdata["author"]
	nome = responsavel.split()[0]
	sobrenome = responsavel.split()[1]
	instituicao = projectdata["institution"]
	funcao = projectdata["role"]
	endereco1 = projectdata["address1"]
	endereco2 = projectdata["address2"]
	cep = projectdata["zip"]
	cidade = projectdata["city"]
	uf = projectdata["uf"]
	fone = projectdata["phone"]
	fax = projectdata["fax"]
	email = projectdata["email"]
	url = projectdata["website"]
	apoio = projectdata["funding"]
	
	#--- Level 1: Identification
	emlBody = (
		"   <title>" + titulo + "</title>\n"
		"   <creator id=\"" + sobrenome.lower() + '.' + nome.lower() + "\">\n"
		"      <individualName>\n"
		"         <givenName>" + nome + "</givenName>\n"
		"         <surName>" + sobrenome + "</surName>\n"
		"      </individualName>\n"
		"      <organizationName>" + instituicao + "</organizationName>\n"
		"      <address>\n"
		"         <deliveryPoint>" + endereco1 + "</deliveryPoint>\n"          
		"         <deliveryPoint>" + endereco2 + "</deliveryPoint>\n"          
		"         <city>" + cidade + "</city>\n"                               
		"         <administrativeArea>" + uf + "</administrativeArea>\n"       
		"         <postalCode>" + cep + "</postalCode>\n"                      
		"         <country>" + pais + "</country>\n"
		"      </address>\n"
		"      <phone phonetype=\"voice\">" + fone + "</phone>\n"              
		"      <phone phonetype=\"fax\">" + fax + "</phone>\n"                 
		"      <electronicMailAddress>" + email + "</electronicMailAddress>\n" 
		"   </creator>\n"
		"   <pubDate>" + str(now.year) + "</pubDate>\n"
		"   <abstract>\n"
		"         <para>" + descricao + "</para>\n"                            
		"   </abstract>\n"
	)
	xmldoc += emlBody
	
	#outfile.write("   <keywordSet>\n")
	#for k in range(len(keywords)):
	#	outfile.write("         <keyword>" + keyword[k] + "</keyword>\n")
	#outfile.write("   </keywordSet>\n")
		
	db_cursor.execute("SELECT MIN(" + headers[LATITUDE] + "), MAX("  + headers[LATITUDE] + "), MIN(" + \
					headers[LONGITUDE] + "), MAX(" + headers[LONGITUDE] + ") FROM Temp")
	latlon = db_cursor.fetchall()
	for row in latlon:
		minlat = row[0]
		maxlat = row[1]
		minlon = row[2]
		maxlon = row[3]
	if minlat is None: minlat = 0.0
	if maxlat is None: maxlat = 0.0
	if minlon is None: minlon = 0.0
	if maxlon is None: maxlon = 0.0
		
	db_cursor.execute("SELECT " + quote_identifier(headers[ELEVATION], "ignore") + " FROM Temp")
	altitude = db_cursor.fetchall()
	minalt = 0
	maxalt = 0
	for row in altitude:
		if row[0] is not None:
			if '-' in str(row[0]):
				alt = to_int(row[0].split('-')[1])
			else:
				alt = to_int(row[0])
		else:	
			alt = 0
		if alt < minalt:
			minalt = alt
		if alt > maxalt:
			maxalt = alt
	unit = get_unit(headers[ELEVATION]) + '.'
	levels += 1
	
	#--- Level 2: Discovery
	emlBody = (
		"   <coverage>\n"
		"       <geographicCoverage>\n"
		"           <geographicDescription>" + ','.join([pais, estado, municipio, localidade]) + "</geographicDescription>\n"
		"           <boundingCoordinates>\n"
		"               <westBoundingCoordinate>" + str(maxlon) + "</westBoundingCoordinate>\n"
		"               <eastBoundingCoordinate>" + str(minlon) + "</eastBoundingCoordinate>\n"
		"               <northBoundingCoordinate>" + str(maxlat) + "</northBoundingCoordinate>\n"
		"               <southBoundingCoordinate>" + str(minlat) + "</southBoundingCoordinate>\n"
		"               <boundingAltitudes>\n"
		"                   <altitudeMinimum>" + str(minalt) + "</altitudeMinimum>\n"
		"                   <altitudeMaximum>" + str(maxalt) + "</altitudeMaximum>\n"
		"                   <altitudeUnits>" + unit + "</altitudeUnits>\n"
		"               </boundingAltitudes>\n"
		"           </boundingCoordinates>\n"
		"       </geographicCoverage>\n"
	)
	xmldoc += emlBody

	db_cursor.execute("SELECT " + headers[DATE] + " FROM Temp WHERE " + headers[DATE] + \
			" IS NOT NULL ORDER BY SUBSTR(" + headers[DATE] + ",7)")
	db_data = db_cursor.fetchall()
	date1 = db_data[0][0]
	date1 = date1.split('-')
	date2 = db_data[len(db_data)-1][0]
	date2 = date2.split('-')

	emlBody = (
		"       <temporalCoverage>\n"
		"          <rangeOfDates>\n"
		"             <beginDate>\n"
		"                <calendarDate>" + '-'.join([date1[2], date1[1], date1[0]]) + "</calendarDate>\n"
		"             </beginDate>\n"
		"             <endDate>\n"
		"                <calendarDate>" + '-'.join([date2[2], date2[1], date2[0]]) + "</calendarDate>\n"
		"             </endDate>\n"
		"          </rangeOfDates>\n"
		"       </temporalCoverage>\n"
	)
	xmldoc += emlBody

	db_cursor.execute("SELECT DISTINCT TRIM(" + headers[FAMILY] + "), TRIM(" + headers[SPECIES] + ")" + \
			" FROM Temp ORDER BY TRIM(" +  headers[FAMILY] + "), TRIM(" + headers[SPECIES] + ")")
	db_data = db_cursor.fetchall()

	emlBody = (
		"       <taxonomicCoverage>\n"
		"          <generalTaxonomicCoverage>" + descricao + "</generalTaxonomicCoverage>\n" 
	)
	xmldoc += emlBody

	grpCheck = ""
	grpEval  = ""
	subCheck = ""
	subEval  = ""
	subSubCheck = ""
	subSubEval  = ""
		
	for row in db_data:
		grpCheck = str(row[0])
		if grpEval <> grpCheck:
			grpEval = grpCheck
			emlBody = (
				"                 <taxonomicClassification>\n" + \
				"                    <taxonRankName>Family</taxonRankName>\n" + \
				"                    <taxonRankValue>" + grpCheck + "</taxonRankValue>\n" + \
				"                 </taxonomicClassification>\n"
			)
			xmldoc += emlBody
			
		subCheck = str(row[1]).split()[0]
		if subEval <> subCheck:
			subEval = subCheck
			emlBody = (
				"                     <taxonomicClassification>\n" + \
				"                        <taxonRankName>Genus</taxonRankName>\n" + \
				"                        <taxonRankValue>" + subCheck + "</taxonRankValue>\n" + \
				"                     </taxonomicClassification>\n"
			)
			xmldoc += emlBody
	
		emlBody = (
			"	                      <taxonomicClassification>\n" + \
			"                            <taxonRankName>Species</taxonRankName>\n" + \
			"                            <taxonRankValue>" + str(row[1]) + "</taxonRankValue>\n" + \
			"                         </taxonomicClassification>\n"
		)
		xmldoc += emlBody

	emlBody = (
		"	   </taxonomicCoverage>\n"
		"   </coverage>\n")
	xmldoc += emlBody
	
	emlBody = (
		"   <maintenance>\n"
		"      <description>\n"
		"         <para>Last access: " + time.ctime(os.path.getatime(projectdata["dataset1"])) + "</para>\n" 
		"         <para>Last modified: " + time.ctime(os.path.getmtime(projectdata["dataset1"])) + "</para>\n"
		"      </description>\n"
		"   </maintenance>\n"
		"   <contact>\n"
		"      <references>" + sobrenome.lower() + '.' + nome.lower() + "</references>\n"
		"   </contact>\n"
		"   <publisher>\n"
		"      <references>" + sobrenome.lower() + '.' +nome.lower() + "</references>\n"
		"   </publisher>\n"
	)
	xmldoc += emlBody
	levels += 1
	
	#--- Level 3: Evaluation
	metodo = projectdata["method"]

	emlBody = (
		"    <methods>\n"
		"       <methodStep>\n"
		"           <description>\n"
		"              <para>\n"
		"                 <literalLayout>" + metodo + "</literalLayout></para>\n"
		"           </description>\n"
		"       </methodStep>\n"      
		"       <sampling>\n"
		"           <studyExtent>\n"
		"              <coverage>\n"
		"                 <temporalCoverage>\n"
		"                     <rangeOfDates>\n"
		"                         <beginDate>\n"
		"                            <calendarDate>" + '-'.join([date1[2], date1[1], date1[0]]) + "</calendarDate>\n"
		"                         </beginDate>\n"
		"                         <endDate>\n"
		"                             <calendarDate>" + '-'.join([date2[2], date2[1], date2[0]]) + "</calendarDate>\n"
		"                         </endDate>\n"
		"                     </rangeOfDates>\n"
		"                 </temporalCoverage>\n"
		"                 <geographicCoverage>\n"
		"                     <geographicDescription>" + ','.join([pais, estado, municipio, localidade]) + "</geographicDescription>\n"
		"                         <boundingCoordinates>\n"
		"                             <westBoundingCoordinate>" + str(maxlon) + "</westBoundingCoordinate>\n"
		"                             <eastBoundingCoordinate>" + str(minlon) + "</eastBoundingCoordinate>\n"
		"                             <northBoundingCoordinate>" + str(maxlat) + "</northBoundingCoordinate>\n"
		"                             <southBoundingCoordinate>" + str(minlat) + "</southBoundingCoordinate>\n"
		"                             <boundingAltitudes>\n"
		"                                 <altitudeMinimum>" + str(minalt) + "</altitudeMinimum>\n"
		"                                 <altitudeMaximum>" + str(maxalt) + "</altitudeMaximum>\n"
		"                                 <altitudeUnits>" + unit + "</altitudeUnits>\n"
		"                             </boundingAltitudes>\n"
		"                         </boundingCoordinates>\n"
		"                 </geographicCoverage>\n"
		"              </coverage>\n"
		"           </studyExtent>\n"
		"           <samplingDescription>\n"
		"               <para>" + metodo + "</para>\n"
		"           </samplingDescription>\n"
		"          <spatialSamplingUnits>\n"
	)              
	xmldoc += emlBody

	db_cursor.execute("SELECT DISTINCT TRIM(" + headers[SAMPLE] + "), " + headers[LATITUDE] + ", " + \
			headers[LONGITUDE] + " FROM Temp ORDER BY TRIM(" + headers[SAMPLE] + ")")
	latlon = db_cursor.fetchall()

	for row in latlon:
		emlBody = (
			"               <coverage>\n"
			"                  <geographicDescription>" + str(row[0]) + "</geographicDescription>\n"
			"                  <boundingCoordinates>\n"
			"                     <westBoundingCoordinate>" + str(row[2]) + "</westBoundingCoordinate>\n"
			"                     <eastBoundingCoordinate>" + str(row[2]) + "</eastBoundingCoordinate>\n"
			"                     <northBoundingCoordinate>" + str(row[1]) + "</northBoundingCoordinate>\n"
			"                     <southBoundingCoordinate>" + str(row[1]) + "</southBoundingCoordinate>\n"
			"                  </boundingCoordinates>\n"
			"               </coverage>\n"
			)
	xmldoc += emlBody

	emlBody = (
		"          </spatialSamplingUnits>\n"
		"       </sampling>\n"
		"    </methods>\n"
	)
	xmldoc += emlBody

	emlBody = (
		"    <project>\n"
		"       <title>" + titulo + "</title>\n"
		"       <personnel>\n"
		"          <references>" + sobrenome.lower() + '.' + nome.lower() + "</references>\n"
		"          <role>" + funcao + "</role>\n"                     
		"       </personnel>\n"
		"       <abstract>\n"
		"          <para>" + descricao + "</para>\n"
		"       </abstract>\n"
		"       <funding>\n"
		"          <para>" + apoio + "</para>\n"
		"       </funding>\n"
		"    </project>\n"
	)
	xmldoc += emlBody
	levels += 1
	
	#--- Level 4: Access
	#emlBody = (
	#	"    <access>\n"
	#	"       <allow>\n"
	#	"         <principal>public</principal>\n" ##--- Acesso (Publico|Restrito)
	#	"         <permission>read</permission>\n" ##--- Permissao (Leitura|Escrita)
	#	"       </allow>\n"
	#	"    </access>\n"
	#	)
	#outfile.write(emlBody)

	#--- Level 5: Integration
	#emlBody = (
	#	)
	#outfile.write(emlBody)

	emlFooter = "</dataset>\n</eml:eml>"
	xmldoc += emlFooter

	outfile = open(unicode(filename), 'w')
	outfile.write(xmldoc.toUtf8())

	db_cursor.close()
	db_connection.close()
	outfile.close()
	return levels

def toKML(filename, projectdata, headers, types, data):
	sqlCreateStr = "CREATE TABLE Temp("
	for i in range(len(headers)):
		if types[i] == "text":
			sqlCreateStr += quote_identifier(headers[i], "replace") + " TEXT"
		elif types[i] == "numeric":
			sqlCreateStr += quote_identifier(headers[i], "replace") + " NUMERIC"
		if i < len(headers) - 1: sqlCreateStr += ", "
	sqlCreateStr += ")"

	sqlInsertStr = "INSERT INTO Temp VALUES("
	for i in range(len(headers)):
		sqlInsertStr += "?"
		if i < len(headers) - 1: sqlInsertStr += ", "
	sqlInsertStr += ")"
		
	db_data = tuple(data)
	db_connection = sqlite3.connect(":memory:")
	db_cursor = db_connection.cursor()
	db_cursor.execute(sqlCreateStr)
	db_cursor.executemany(sqlInsertStr, db_data)
	
	db_cursor.execute("SELECT TRIM(" + headers[SAMPLE] + "), " + headers[LATITUDE] + ", " + \
					headers[LONGITUDE] + ", " + quote_identifier(headers[ELEVATION], "ignore") + \
					", TRIM(" + headers[OBS] + ") FROM Temp")
	db_data = db_cursor.fetchall()
	
	kmlHeader = "<?xml version=\"1.0\" encoding=\"UTF-8\"?>\n<kml xmlns=\"http://www.opengis.net/kml/2.2\">\n"
	kmlBody = (
		"<Folder>\n"
		"   <Style id=\"normalPlaceMarker\">\n"
		"   <IconStyle>\n"
		"   <Icon>\n"
		"      <href>http://maps.google.com/mapfiles/kml/pal3/icon38.png</href>\n"
		"   </Icon>\n"
		"   </IconStyle>\n"
		"   </Style>\n"
		)
	kmlFooter = "</Folder>\n</kml>"

	count = 0
	for row in db_data:
		local = str(row[0])
		latitude = iif(row[1] != None, to_float(row[1]), 0.0000)
		longitude = iif(row[2] != None, to_float(row[2]), 0.0000)
		if row[3] is not None:
			if '-' in str(row[3]):
				altitude = to_int(row[3].split('-')[1])
			else:
				altitude = to_int(row[3])
		else:	
			altitude = 0.0
		desc = str(row[4])
		kml = (
			"   <Placemark>\n"
			"      <name>" + local + "</name>\n"
			"      <description>" + desc + "</description>\n"
			"      <Point>\n"
			"          <coordinates>%f,%f,%d</coordinates>\n"
			"      </Point>\n"
			"   </Placemark>\n"
		) %(longitude, latitude, altitude)
		kmlBody += kml
		count += 1
	
	kmlOutput = kmlHeader + kmlBody + kmlFooter
	outfile = open(unicode(filename), 'w')
	outfile.write(kmlOutput)
	
	db_cursor.close()
	db_connection.close()
	outfile.close()
	return count

def toRDE(filename, projectdata, headers, types, data):
	sqlCreateStr = "CREATE TABLE Temp("
	for i in range(len(headers)):
		if types[i] == "text":
			sqlCreateStr += quote_identifier(headers[i], "replace") + " TEXT"
		elif types[i] == "numeric":
			sqlCreateStr += quote_identifier(headers[i], "replace") + " NUMERIC"
		if i < len(headers) - 1: sqlCreateStr += ", "
	sqlCreateStr += ")"

	sqlInsertStr = "INSERT INTO Temp VALUES("
	for i in range(len(headers)):
		sqlInsertStr += "?"
		if i < len(headers) - 1: sqlInsertStr += ", "
	sqlInsertStr += ")"
		
	db_data = tuple(data)
	db_connection = sqlite3.connect(":memory:")
	db_cursor = db_connection.cursor()
	db_cursor.execute(sqlCreateStr)
	db_cursor.executemany(sqlInsertStr, db_data)
	
	country = projectdata["country"].encode("utf-8")
	majorarea = projectdata["state"].encode("utf-8")
	minorarea = projectdata["province"].encode("utf-8")
	gazetteer = projectdata["locality"].encode("utf-8")
	
	db_cursor.execute("SELECT " + quote_identifier(headers[ELEVATION], "ignore") + " FROM Temp")
	altitude = db_cursor.fetchall()
	altmax = 0
	for row in altitude:
		if row[0] is not None:
			if '-' in str(row[0]):
				alt = to_int(row[0].split('-')[1])
			else:
				alt = to_int(row[0])
		else:	
			alt = 0
		if alt > altmax:
			altmax = alt
	unit = get_unit(headers[ELEVATION])
	
	db_cursor.execute("SELECT TRIM(" + headers[FAMILY] + ", TRIM(" + headers[SPECIES] + "), TRIM(" + \
		headers[COLLECTOR] + "), " + headers[NUMCOL] + ", " + headers[DATE] + ", TRIM(" + headers[OBS] + "), TRIM(" + \
		headers[LOCALITY] + "), " + quote_identifier(headers[ELEVATION], "ignore") + ", " + \
		headers[LATITUDE] + ", " + headers[LONGITUDE] + " FROM Temp")
	all_rows = db_cursor.fetchall()

	dbfile = dbf.Dbf(str(filename), new=True)
	dbfile.addField(
		("TAG","C",1),
		("DEL","C",1),
		("BOTRECCAT","C",1),
		("RDEIMAGES","C",128),
		("DUPS","C",40),
		("BARCODE","C",15),
		("ACCESSION","C",15),
		("COLLECTOR","C",80),
		("ADDCOLL","C",80),
		("PREFIX","C",10),
		("NUMBER","C",15),
		("SUFFIX","C",6),
		("COLLDD","N",2,0),
		("COLLMM","N",2,0),
		("COLLYY","N",4,0),
		("DATERES","C",5),
		("DATETEXT","C",128),
		("FAMILY","C",30),
		("GENUS","C",25),
		("SP1","C",25),
		("AUTHOR1","C",75),
		("RANK1","C",10),
		("SP2","C",25),
		("AUTHOR2","C",75),
		("RANK2","C",10),
		("SP3","C",25),
		("AUTHOR3","C",75),
		("UNIQUE","N",1,0),
		("PLANTDESC","C",128),
		("PHENOLOGY","C",10),
		("DETBY","C",30),
		("DETDD","N",2,0),
		("DETMM","N",2,0),
		("DETYY","N",4,0),
		("DETSTATUS","C",5),
		("DETNOTES","C",128),
		("COUNTRY","C",50),
		("MAJORAREA","C",30),
		("MINORAREA","C",30),
		("GAZETTEER","C",50),
		("LOCNOTES","C",128),
		("HABITATTXT","C",128),
		("CULTIVATED","L",1),
		("CULTNOTES","C",128),
		("ORIGINSTAT","C",5),
		("ORIGINID","C",15),
		("ORIGINDB","C",10),
		("NOTES","C",128),
		("LAT","N",16,10),
		("NS","C",1),
		("LONG","N",16,10),
		("EW","C",1),
		("LLGAZ","C",5),
		("LLUNIT","C",5),
		("LLRES","C",5),
		("LLORIG","C",20),
		("LLDATUM","C",10),
		("QDS","C",10),
		("ALT","C",8),
		("ALTMAX","C",8),
		("ALTUNIT","C",1),
		("ALTRES","C",5),
		("ALTTEXT","C",128),
		("INITIAL","N",2,0),
		("AVAILABLE","N",2,0),
		("CURATENOTE","C",128),
		("NOTONLINE","C",1),
		("LABELTOTAL","N",2,0),
		("VERNACULAR","C",40),
		("LANGUAGE","C",40),
		("GEODATA","C",128),
		("LATLONG","C",50),
		("COLLECTED","C",20),
		("MONTHNAME","C",10),
		("DETBYDATE","C",20),
		("CHECKWHO","C",5),
		("CHECKDATE","D",8),
		("CHECKNOTE","C",128),
		("UNIQUEID","C",15),
		("RDESPEC","C",128)
	)

	dt = datetime.now()
	reccount = 0
	for row in all_rows:
		collector = str(row[2])
		number = str(row[3])
		date = iif(row[4] != None, str(row[4]), "00-00-0000")
		colldd = to_int(date.split('-')[0])
		collmm = to_int(date.split('-')[1])
		collyy = to_int(date.split('-')[2])
		family = str(row[0])
		species = str(row[1])
		genus, cf, sp1, author1, sp2, infraname, author2 = parse_name(species)
		locnotes = str(row[6])
		notes = str(row[5])
		lat = to_float(row[8])
		ns = iif(lat < 0, 'S', 'N')
		long = to_float(row[9])
		ew = iif(long < 0, 'E', 'W')
		if row[7] is not None:
			if '-' in str(row[7]):
				alt = to_int(str(row[7]).split('-')[0])
			else:
				alt = to_int(row[7])
		else:	
			alt = 0
			
		rec = dbfile.newRecord()
		rec["TAG"] = ""
		rec["DEL"] = ""
		rec["BOTRECCAT"] = ""
		rec["RDEIMAGES"] = ""
		rec["DUPS"] = ""
		rec["BARCODE"] = ""
		rec["ACCESSION"] = ""
		rec["COLLECTOR"] = collector
		rec["ADDCOLL"] = ""
		rec["PREFIX"] = ""
		rec["NUMBER"] = number
		rec["SUFFIX"] = ""
		rec["COLLDD"] = colldd
		rec["COLLMM"] = collmm
		rec["COLLYY"] = collyy
		rec["DATERES"] = ""
		rec["DATETEXT"] = ""
		rec["FAMILY"] = family
		rec["GENUS"] = genus
		rec["SP1"] = sp1
		rec["AUTHOR1"] = author1
		rec["RANK1"] = sp2
		rec["SP2"] = infraname
		rec["AUTHOR2"] = author2
		rec["RANK2"] = ""
		rec["SP3"] = ""
		rec["AUTHOR3"] = ""
		rec["UNIQUE"] = 0
		rec["PLANTDESC"] = ""
		rec["PHENOLOGY"] = ""
		rec["DETBY"] = ""
		rec["DETDD"] = 0
		rec["DETMM"] = 0
		rec["DETYY"] = 0
		rec["DETSTATUS"] = ""
		rec["DETNOTES"] = ""
		rec["COUNTRY"] = country
		rec["MAJORAREA"] = majorarea
		rec["MINORAREA"] = minorarea
		rec["GAZETTEER"] = gazetteer
		rec["LOCNOTES"] = locnotes
		rec["HABITATTXT"] = ""
		rec["CULTIVATED"] = ""
		rec["CULTNOTES"] = ""
		rec["ORIGINSTAT"] = ""
		rec["ORIGINID"] = ""
		rec["ORIGINDB"] = ""
		rec["NOTES"] = notes
		rec["LAT"] = lat
		rec["NS"] = ns
		rec["LONG"] = long
		rec["EW"] = ew
		rec["LLGAZ"] = ""
		rec["LLUNIT"] = ""
		rec["LLRES"] = ""
		rec["LLORIG"] = ""
		rec["LLDATUM"] = ""
		rec["QDS"] = ""
		rec["ALT"] = alt
		rec["ALTMAX"] = altmax
		rec["ALTUNIT"] = unit
		rec["ALTRES"] = ""
		rec["ALTTEXT"] = ""
		rec["INITIAL"] = 0
		rec["AVAILABLE"] = 0
		rec["CURATENOTE"] = ""
		rec["NOTONLINE"] = ""
		rec["LABELTOTAL"] = 0
		rec["VERNACULAR"] = ""
		rec["LANGUAGE"] = ""
		rec["GEODATA"] = ""
		rec["LATLONG"] = ""
		rec["COLLECTED"] = ""
		rec["MONTHNAME"] = ""
		rec["DETBYDATE"] = ""
		rec["CHECKWHO"] = ""
		rec["CHECKDATE"] = dt
		rec["CHECKNOTE"] = ""
		rec["UNIQUEID"] = ""
		rec["RDESPEC"] = ""
		rec.store()
		reccount += 1

	db_cursor.close()
	db_connection.close()
	dbfile.close()	
	return reccount

def toCEP(filename, projectdata, headers, types, data):
	sqlCreateStr = "CREATE TABLE Temp("
	for i in range(len(headers)):
		if types[i] == "text":
			sqlCreateStr += quote_identifier(headers[i], "replace") + " TEXT"
		elif types[i] == "numeric":
			sqlCreateStr += quote_identifier(headers[i], "replace") + " NUMERIC"
		if i < len(headers) - 1: sqlCreateStr += ", "
	sqlCreateStr += ")"

	sqlInsertStr = "INSERT INTO Temp VALUES("
	for i in range(len(headers)):
		sqlInsertStr += "?"
		if i < len(headers) - 1: sqlInsertStr += ", "
	sqlInsertStr += ")"
		
	db_data = tuple(data)
	db_connection = sqlite3.connect(":memory:")
	db_cursor = db_connection.cursor()
	db_cursor.execute(sqlCreateStr)
	db_cursor.executemany(sqlInsertStr, db_data)

	db_cursor.execute("SELECT DISTINCT TRIM(" + headers[SAMPLE] + ") FROM Temp")
	amostras = db_cursor.fetchall()
	rows = len(amostras)

	db_cursor.execute("SELECT DISTINCT TRIM(" + headers[SPECIES] + ") FROM Temp")
	especies = db_cursor.fetchall()
	cols = len(especies)
	
	outfile = open(unicode(filename), "w")

	title = projectdata["title"]
	outfile.write(title + '\n')
	
	format = "(I4,2X,7(I5,F5.1))\n"
	outfile.write(format)

	count = 1
	for i in range(rows):
		amostra = str(amostras[i][0])
		outfile.write("%4d  " % count)
		count += 1
		for j in range(cols):
			species = especies[j][0]
			numero = j + 1
			db_cursor.execute("SELECT COUNT(TRIM(" + headers[SPECIES] + ")) FROM Temp WHERE " + \
						headers[SPECIES] + " = '" + species + "' AND " + \
						headers[SAMPLE] + " = '" + amostra + "'")
			freq = db_cursor.fetchone()[0]
			if freq == None:
				freq = 0
			if freq > 0:
				outfile.write("%5d%5.1f" % (numero, freq))
		outfile.write('\n')
	outfile.write("0\n") 

	for j in range(cols):
		genus, cf, species, author1, subsp, infraname, author2 = parse_name(especies[j][0])
		outfile.write(genus[:4] + species[:4])
	outfile.write('\n')

	for i in range(rows):
		sample = amostras[i][0]
		outfile.write(sample.ljust(8))

	outfile.close()		
	db_cursor.close()
	db_connection.close()
	return rows, cols

def toMatrix(filename, projectdata, headers, types, data, format, kind):
	sqlCreateStr = "CREATE TABLE Temp("
	for i in range(len(headers)):
		if types[i] == "text":
			sqlCreateStr += quote_identifier(headers[i], "replace") + " TEXT"
		elif types[i] == "numeric":
			sqlCreateStr += quote_identifier(headers[i], "replace") + " NUMERIC"
		if i < len(headers) - 1: sqlCreateStr += ", "
	sqlCreateStr += ")"

	sqlInsertStr = "INSERT INTO Temp VALUES("
	for i in range(len(headers)):
		sqlInsertStr += "?"
		if i < len(headers) - 1: sqlInsertStr += ", "
	sqlInsertStr += ")"
		
	db_data = tuple(data)
	db_connection = sqlite3.connect(":memory:")
	db_cursor = db_connection.cursor()
	db_cursor.execute(sqlCreateStr)
	db_cursor.executemany(sqlInsertStr, db_data)

	db_cursor.execute("SELECT DISTINCT TRIM(" + headers[SAMPLE] + ") FROM Temp")
	amostras = db_cursor.fetchall()
	rows = len(amostras)

	db_cursor.execute("SELECT DISTINCT TRIM(" + headers[SPECIES] + ") FROM Temp")
	especies = db_cursor.fetchall()
	cols = len(especies)

	outfile = open(unicode(filename), "w")

	title = projectdata["title"]
	if format == "mvsp":
		outfile.write("*MVSP3 " + str(rows) + ' ' + str(cols) + ' ' + title + '\n')
	elif format == "ntsys":
		outfile.write("1 " + str(rows) + "L " + str(cols) + "L " + " 0\n")
	
	if format == "ntsys":
		for i in range(rows):
			sample = amostras[i][0].replace(' ', '')
			outfile.write(sample[:8] + ' ')
		outfile.write('\n')

	for j in range(cols):
		genus, cf, species, author1, subsp, infraname, author2 = parse_name(especies[j][0])
		outfile.write(genus[:4] + species[:4] + ' ')
	outfile.write('\n')
	
	for i in range(rows):
		sample = str(amostras[i][0])
		if format <> "ntsys":
			outfile.write(sample.replace(' ', '') + ' ')
		for j in range(cols):
			species = especies[j][0]
			if kind == 1: 
				db_cursor.execute("SELECT TRIM(" + headers[SPECIES] + ") FROM Temp WHERE TRIM(" + \
						headers[SPECIES] + ") = '" + species + "' AND TRIM(" + \
						headers[SAMPLE] + ") = '" + sample + "'")
				found = db_cursor.fetchone() != None
				if found:
					bin = 1
				else:
					bin = 0
				outfile.write("%d" % bin)
			elif kind == 2:	
				db_cursor.execute("SELECT COUNT(TRIM(" + headers[SPECIES] + ")) FROM Temp WHERE TRIM(" + \
						headers[SPECIES] + ") = '" + species + "' AND TRIM(" + \
						headers[SAMPLE] + ") = '" + sample + "'")
				freq = db_cursor.fetchone()[0]
				if freq == None:
					freq = 0
				outfile.write("%d" % freq)
			if j < cols - 1: outfile.write(' ')
		outfile.write('\n')
	
	outfile.close()
	db_cursor.close()
	db_connection.close()
	return rows, cols

def toCSV(filename, projectdata, headers, types, data, kind):
	sqlCreateStr = "CREATE TABLE Temp("
	for i in range(len(headers)):
		if types[i] == "text":
			sqlCreateStr += quote_identifier(headers[i], "replace") + " TEXT"
		elif types[i] == "numeric":
			sqlCreateStr += quote_identifier(headers[i], "replace") + " NUMERIC"
		if i < len(headers) - 1: sqlCreateStr += ", "
	sqlCreateStr += ")"

	sqlInsertStr = "INSERT INTO Temp VALUES("
	for i in range(len(headers)):
		sqlInsertStr += "?"
		if i < len(headers) - 1: sqlInsertStr += ", "
	sqlInsertStr += ")"
		
	db_data = tuple(data)
	db_connection = sqlite3.connect(":memory:")
	db_cursor = db_connection.cursor()
	db_cursor.execute(sqlCreateStr)
	db_cursor.executemany(sqlInsertStr, db_data)

	db_cursor.execute("SELECT DISTINCT TRIM(" + headers[SAMPLE] + ") FROM Temp")
	amostras = db_cursor.fetchall()
	rows = len(amostras)

	db_cursor.execute("SELECT DISTINCT TRIM(" + headers[SPECIES] + ") FROM Temp")
	especies = db_cursor.fetchall()
	cols = len(especies)

	outfile = open(unicode(filename), "w")

	for j in range(cols):
		genus, cf, species, author1, subsp, infraname, author2 = parse_name(especies[j][0])
		outfile.write(genus[:4] + species[:4].replace('.', ''))
		if j < cols - 1: outfile.write(',') 
	outfile.write('\n')
	
	for i in range(rows):
		sample = str(amostras[i][0])
		outfile.write(sample + ',')
		for j in range(cols):
			species = especies[j][0]
			if kind == 1: 
				db_cursor.execute("SELECT TRIM(" + headers[SPECIES] + ") FROM Temp WHERE TRIM(" + \
						headers[SPECIES] + ") = '" + species + "' AND TRIM(" + \
						headers[SAMPLE] + ") = '" + sample + "'")
				found = db_cursor.fetchone() != None
				if found: 
					bin = 1
				else:
					bin = 0
				outfile.write("%d" % bin)
			elif kind == 2:	
				db_cursor.execute("SELECT COUNT(TRIM(" + headers[SPECIES] + ")) FROM Temp WHERE TRIM(" + \
						headers[SPECIES] + ") = '" + species + "' AND TRIM(" + \
						headers[SAMPLE] + ") = '" + sample + "'")
				freq = db_cursor.fetchone()[0]
				if freq == None:
					freq = 0
				outfile.write("%d" % freq)
			if j < cols - 1: outfile.write(',')
		outfile.write('\n')
	
	outfile.close()
	db_cursor.close()
	db_connection.close()
	return rows, cols

def toFitopac1(filename, projectdata, headers, types, data, tipo):
	vocabulary = ["ALTURA","HEIGHT","X","Y","DISTANCIA","DISTANCE",
		"PAP","P.A.P.","CAP","C.A.P.","DAP","D.A.P.","DIAMETRO","DIAMETER",
		"PERIMETRO","PERIMETER","DBH","PBH"]
	
	sqlCreateStr = "CREATE TABLE Temp("
	for i in range(len(headers)):
		if types[i] == "text":
			sqlCreateStr += quote_identifier(headers[i], "replace") + " TEXT"
		elif types[i] == "numeric":
			sqlCreateStr += quote_identifier(headers[i], "replace") + " NUMERIC"
		if i < len(headers) - 1: sqlCreateStr += ", "
	sqlCreateStr += ")"

	sqlInsertStr = "INSERT INTO Temp VALUES("
	for i in range(len(headers)):
		sqlInsertStr += "?"
		if i < len(headers) - 1: sqlInsertStr += ", "
	sqlInsertStr += ")"
		
	db_data = tuple(data)
	db_connection = sqlite3.connect(":memory:")
	db_cursor = db_connection.cursor()
	db_cursor.execute(sqlCreateStr)
	db_cursor.executemany(sqlInsertStr, db_data)
	
	fname1 = os.path.splitext(str(filename))[0] + ".nms"
	outfile = open(unicode(fname1), "w")
	
	db_cursor.execute("SELECT DISTINCT TRIM(" + headers[FAMILY] + ") FROM Temp ORDER BY TRIM(" + headers[FAMILY] + ")")
	db_data = db_cursor.fetchall()

	f = 0
	vetfam = []
	for row in db_data:
		vfamilia = str(row[0])
		vetfam.append(vfamilia)
		outfile.write(str(f + 1) + ' ' + vetfam[f] + '\n')
		f += 1
	outfile.write("999\n")

	db_cursor.execute("SELECT DISTINCT TRIM(" + headers[FAMILY] + "), TRIM(" + headers[SPECIES] + ")" + \
					" FROM Temp ORDER BY TRIM(" + headers[SPECIES] + ")")
	db_data = db_cursor.fetchall()

	s = 0
	vetspp = []
	vetnum = []
	for row in db_data:
		vfam = str(row[0])
		pos = vetfam.index(vfam) + 1
		vetnum.append(pos)
		vspp = str(row[1])
		vetspp.append(vspp)
		outfile.write(str(s + 1) + ' ' + str(vetnum[s]) + ' ' + vetspp[s] + '\n')
		s += 1
	outfile.close()

	fname2 = os.path.splitext(str(filename))[0] + ".dad"
	outfile = open(unicode(fname2), "w")
	
	db_cursor.execute("SELECT TRIM(" + headers[SAMPLE] + "), " + headers[INDIVIDUAL] + ", TRIM(" + headers[SPECIES] + ")" + \
					" FROM Temp ORDER BY TRIM(" + headers[SAMPLE] + "), " + headers[INDIVIDUAL])
	db_data = db_cursor.fetchall()
	
	fields = []
	for pos in range(len(headers)):
		descriptor = unicode_to_ascii(headers[pos]).upper()
		if descriptor.split(' ')[0] in vocabulary:
			fields.append(pos)

	i = k = l = 0
	aux = ""
	for row in db_data:
		if tipo == 2:   #--- contador de quadrantes
			l += 1
			if l >= 4:
				l = 0
				k += 1
		if tipo == 1:    #--- Parcela
			vloc = str(row[0])
			if aux <> vloc:
				aux = vloc
				outfile.write(vloc.replace(' ', '') + '\n')
			outfile.write(str(i + 1))
			i += 1
		elif tipo == 2:  #--- Quadrante
			outfile.write(str(k + 1) + ' ' + str(i + 1))
			i += 1
	
		for pos in range(len(fields)):
			db_cursor.execute("SELECT " + quote_identifier(headers[fields[pos]], "replace") + \
			" FROM Temp WHERE " + headers[INDIVIDUAL] + " = '" + str(row[1]) + "'")
			values = db_cursor.fetchall()
			for value in values:	
				outfile.write(' ' + str(value[0]))
	
		mspp = str(row[2])
		pos = vetspp.index(mspp) + 1
		outfile.write(' ' + str(pos) + '\n')
	
	outfile.close()
	db_cursor.close()
	db_connection.close()
	return f, s, i

def toFitopac2(filename, projectdata, headers, types, data):
	vocabulary = ["ALTURA","HEIGHT","X","Y","DISTANCIA","DISTANCE",
		"PAP","P.A.P.","CAP","C.A.P.","DAP","D.A.P.","DIAMETRO","DIAMETER",
		"PERIMETRO","PERIMETER","DBH","PBH"]
		
	sqlCreateStr = "CREATE TABLE Temp("
	for i in range(len(headers)):
		if types[i] == "text":
			sqlCreateStr += quote_identifier(headers[i], "replace") + " TEXT"
		elif types[i] == "numeric":
			sqlCreateStr += quote_identifier(headers[i], "replace") + " NUMERIC"
		if i < len(headers) - 1: sqlCreateStr += ", "
	sqlCreateStr += ")"

	sqlInsertStr = "INSERT INTO Temp VALUES("
	for i in range(len(headers)):
		sqlInsertStr += "?"
		if i < len(headers) - 1: sqlInsertStr += ", "
	sqlInsertStr += ")"
		
	db_data = tuple(data)
	db_connection = sqlite3.connect(":memory:")
	db_cursor = db_connection.cursor()
	db_cursor.execute(sqlCreateStr)
	db_cursor.executemany(sqlInsertStr, db_data)
	
	outfile = open(unicode(filename), "w")
	outfile.write("FPD 2.1\n")
	outfile.write(projectdata["title"].encode("utf-8") + '\n')
	outfile.write(projectdata["author"].encode("utf-8") + '\n')
	outfile.write(str(projectdata["datetime"]) + '\n')
	outfile.write(projectdata["state"].encode("utf-8") + '\n')
	outfile.write(projectdata["province"].encode("utf-8") + '\n')
	outfile.write(projectdata["locality"].strip().encode("utf-8") + '\n')
	latitude = to_float(projectdata["latitude"])
	longitude = to_float(projectdata["longitude"])
	if latitude < 0.0:
		outfile.write("S")
	else:
		outfile.write("N")
	if longitude < 0.0:
		outfile.write("W ")
	else:
		outfile.write("E ")
	vlatdeg, vlatmin, vlatsec = degtodms(latitude)
	vlondeg, vlonmin, vlonsec = degtodms(longitude)
	outfile.write(str(abs(vlatdeg)) + ' ' + str(vlatmin) + ' ' + str(to_int(vlatsec)) + \
				str(abs(vlondeg)) + ' ' + str(vlonmin) + ' ' + str(to_int(vlonsec)) + '\n')
	altitude = to_int(projectdata["elevation"])
	outfile.write(str(altitude) + '\n')
	outfile.write("P\n")
	size = projectdata["size"].lower().replace(' ', '')
	length = to_float(size.split('x')[0])
	width = to_float(strip_letters(size.split('x')[1]))
	
	db_cursor.execute("SELECT DISTINCT TRIM(" + headers[FAMILY] + ") FROM Temp ORDER BY TRIM(" + headers[FAMILY] + ")")
	db_data = db_cursor.fetchall()

	f = 0
	vetfam = []
	for row in db_data:
		vfamilia = str(row[0])
		vetfam.append(vfamilia)
		f += 1
	
	db_cursor.execute("SELECT DISTINCT TRIM(" + headers[SPECIES] + ") FROM Temp ORDER BY TRIM(" + headers[SPECIES] + ")")
	db_data = db_cursor.fetchall()
	
	s = 0
	vetspp = []
	for row in db_data:
		vspp = str(row[0])
		vetspp.append(vspp)
		s += 1
		
	db_cursor.execute("SELECT DISTINCT TRIM(" + headers[SAMPLE] + ") FROM Temp ORDER BY TRIM(" + headers[SAMPLE] + ")")
	db_data = db_cursor.fetchall()
	
	n = 0
	samples = {}
	for row in db_data:
		vsample = str(row[0])
		db_cursor.execute("SELECT COUNT(*) FROM TEMP WHERE TRIM(" + headers[SAMPLE] + ") = '" + vsample + "'")
		count = db_cursor.fetchone()[0]
		samples[vsample] = count
		n += 1
		
	db_cursor.execute("SELECT COUNT(*) FROM Temp")
	ni = db_cursor.fetchone()[0]
	
	outfile.write(str(f) + ' ' + str(s) + ' ' + str(ni) + ' ' + str(n) + ' ' + str(length) + ' ' + str(width) + '\n')
	outfile.write("TRUE TRUE\n")
	outfile.write("TRUE TRUE\n")
	outfile.write("TRUE DC\n") 
	outfile.write("\n\n\n")
	
	for f in range(len(vetfam)):
		outfile.write("T " + vetfam[f] + '\n')
		
	for s in range(len(vetspp)):
		outfile.write("T " + str(s + 1) + ' ' + vetspp[s] + '\n')

	for y in samples:
		outfile.write("T " + str(samples[y]) + ' ' + y.replace(' ', '') + '\n')
		
	db_cursor.execute("SELECT TRIM(" + headers[SAMPLE] + "), " + headers[INDIVIDUAL] + ", TRIM(" + headers[SPECIES] + ")" + \
					" FROM Temp ORDER BY TRIM(" + headers[SAMPLE] + "), " + headers[INDIVIDUAL])
	db_data = db_cursor.fetchall()
	
	fields = []
	for pos in range(len(headers)):
		descriptor = unicode_to_ascii(headers[pos]).upper()
		if descriptor.split(' ')[0] in vocabulary:
			fields.append(pos)
			
	for row in db_data:
		outfile.write(str(row[1]) + ' ')
		values = []
		for pos in range(len(fields)):
			db_cursor.execute("SELECT " + quote_identifier(headers[fields[pos]], "replace") + \
			" FROM Temp WHERE " + headers[INDIVIDUAL] + " = '" + str(row[1]) + "'")
			value = db_cursor.fetchone()[0]
			values.append(value)
		mspp = str(row[2])
		pos = vetspp.index(mspp) + 1
		outfile.write(str(values[0]) + ' ')
		outfile.write(str(pos) + ' ')
		plus = str(values[1]).count('+')
		if plus > 1:
			outfile.write(str(plus + 1) + ' ' + ' '.join(values[1].split('+')))
		else:
			plus = 1
			outfile.write(str(plus) + ' ' + str(values[1]))
		outfile.write('\n')
	
	outfile.close()
	db_cursor.close()
	db_connection.close()
	return f, s, ni, n

def toFPM(filename, projectdata, headers, types, data, kind):
	sqlCreateStr = "CREATE TABLE Temp("
	for i in range(len(headers)):
		if types[i] == "text":
			sqlCreateStr += quote_identifier(headers[i], "replace") + " TEXT"
		elif types[i] == "numeric":
			sqlCreateStr += quote_identifier(headers[i], "replace") + " NUMERIC"
		if i < len(headers) - 1: sqlCreateStr += ", "
	sqlCreateStr += ")"

	sqlInsertStr = "INSERT INTO Temp VALUES("
	for i in range(len(headers)):
		sqlInsertStr += "?"
		if i < len(headers) - 1: sqlInsertStr += ", "
	sqlInsertStr += ")"
		
	db_data = tuple(data)
	db_connection = sqlite3.connect(":memory:")
	db_cursor = db_connection.cursor()
	db_cursor.execute(sqlCreateStr)
	db_cursor.executemany(sqlInsertStr, db_data)

	db_cursor.execute("SELECT DISTINCT TRIM(" + headers[SAMPLE] + ") FROM Temp")
	amostras = db_cursor.fetchall()
	rows = len(amostras)

	db_cursor.execute("SELECT DISTINCT TRIM(" + headers[SPECIES] + ") FROM Temp")
	especies = db_cursor.fetchall()
	cols = len(especies)

	outfile = open(unicode(filename), "w")
	outfile.write("FPM 2.1\n")
	outfile.write(projectdata["title"].encode("utf-8") + '\n')
	outfile.write(projectdata["author"].encode("utf-8") + '\n')
	outfile.write(str(projectdata["datetime"]) + '\n')
	outfile.write(iif(kind == 1, 'P', 'Q') + '\n')
	outfile.write(str(cols) + ' ' + str(rows) + '\n')
	outfile.write("especies\n")
	outfile.write("amostras\n")
	outfile.write("\n\n\n")

	for j in range(cols):
		genus, cf, species, author1, subsp, infraname, author2 = parse_name(especies[j][0])
		outfile.write("T" + iif(kind == 1, 'P', 'Q') + ' ' + genus + ' ' + species + '\n')
	
	for i in range(rows):
		sample = str(amostras[i][0])
		outfile.write("T " + sample.replace(' ', '') + '\n')
	
	for i in range(rows):
		for j in range(cols):
			species = especies[j][0]
			if kind == 1: 
				db_cursor.execute("SELECT TRIM(" + headers[SPECIES] + ") FROM Temp WHERE TRIM(" + \
						headers[SPECIES] + ") = '" + species + "' AND TRIM(" + \
						headers[SAMPLE] + ") = '" + sample + "'")
				found = db_cursor.fetchone() != None
				if found: 
					bin = 1
				else:
					bin = 0
				outfile.write("%d" % bin)
			elif kind == 2:	
				db_cursor.execute("SELECT COUNT(TRIM(" + headers[SPECIES] + ")) FROM Temp WHERE TRIM(" + \
						headers[SPECIES] + ") = '" + species + "' AND TRIM(" + \
						headers[SAMPLE] + ") = '" + sample + "'")
				freq = db_cursor.fetchone()[0]
				if freq == None:
					freq = 0
				outfile.write("%d" % freq)
			if j < cols - 1: outfile.write(' ')
		outfile.write('\n')
	
	outfile.close()
	db_cursor.close()
	db_connection.close()
	return rows, cols

def toSHP(filename, projectdata, headers, types, data):
	sqlCreateStr = "CREATE TABLE Temp("
	for i in range(len(headers)):
		if types[i] == "text":
			sqlCreateStr += quote_identifier(headers[i], "replace") + " TEXT"
		elif types[i] == "numeric":
			sqlCreateStr += quote_identifier(headers[i], "replace") + " NUMERIC"
		if i < len(headers) - 1: sqlCreateStr += ", "
	sqlCreateStr += ")"

	sqlInsertStr = "INSERT INTO Temp VALUES("
	for i in range(len(headers)):
		sqlInsertStr += "?"
		if i < len(headers) - 1: sqlInsertStr += ", "
	sqlInsertStr += ")"
		
	db_data = tuple(data)
	db_connection = sqlite3.connect(":memory:")
	db_cursor = db_connection.cursor()
	db_cursor.execute(sqlCreateStr)
	db_cursor.executemany(sqlInsertStr, db_data)
	
	db_cursor.execute("SELECT TRIM(" + headers[SAMPLE] + "), " + headers[LATITUDE] + ", " + \
					headers[LONGITUDE] + ", " + quote_identifier(headers[ELEVATION], "ignore") + " FROM Temp")
	db_data = db_cursor.fetchall()
	
	shp = shapefile.Writer(shapefile.POINT)
	shp.field("LOCAL", 'C', 65)
	shp.field("LATITUDE", 'F', 11, 6)
	shp.field("LONGITUDE", 'F', 11, 6)
	shp.field("ELEVATION", 'F', 7, 2)
	
	count = 0
	for row in db_data:
		local = str(row[0])
		latitude = iif(row[1] != None, to_float(row[1]), 0.0000)
		longitude = iif(row[2] != None, to_float(row[2]), 0.0000)
		if row[3] is not None:
			if '-' in str(row[3]):
				altitude = to_int(row[3].split('-')[1])
			else:
				altitude = to_int(row[3])
		else:	
			altitude = 0.0
		shp.point(longitude, latitude, 0, 0)
		shp.record(local, latitude, longitude, altitude)
		shp.autoBalance = 1
		count += 1
	
	shp.save(str(filename))
	db_cursor.close()
	db_connection.close()
	return count
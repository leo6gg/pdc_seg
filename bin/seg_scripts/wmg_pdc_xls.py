#!/usr/bin/python
# -*- coding: utf-8 -*-
import os
import re
import xlwt
import xlrd
#import xlwt3 as xlwt
from xml.etree import ElementTree as ET
import ConfigParser
import time
import xlsxwriter
import sys,os.path
import datetime

curPath = os.getcwd()
tmpPath = os.path.abspath(os.path.join(curPath,os.path.pardir))
workPath = os.path.abspath(os.path.join(tmpPath,os.path.pardir))
configPath = curPath+'/config/'
#print configPath
globalFile = configPath + 'global.cfg'
#print ('globalFile == %s' % globalFile)
config = ConfigParser.ConfigParser()
config.readfp(open(globalFile,'rb'))
outputPath = curPath + '/output/seg_data/'
FILEPATH = config.get("global","pmFilePath")
if FILEPATH == '':
    print('need to configure PM file path in '+ configPath + 'global.cfg')
    sys.exit()


CFGPATH = configPath + 'statistics_seg.cfg'
fileNameList = []

valueStr = re.compile('.+measValue.+')

#configuration by user-defined
partLable = ("<",">")
sectionLable = ("{","}")
#endlineLable = "\r\n" # row tag for windows
endlineLable = "\n"   # row tag for linux
equalLable = "=" # equal sign
noteLable = '#' # note sign

# get all contents from configure file ---- map
def getPlatformMap(strtmp,lable1 = partLable,lable2 = sectionLable):
    tmp = strtmp.split(lable1[0])
    tmp = [elem for elem in tmp if len(elem) > 1]
    tmp = [elem for elem in tmp if elem.rfind(lable1[1]) > 0]
    platdict = {}
    for elem in tmp:
        key = elem[0:elem.find(lable1[1]):]
        value = elem[elem.find(lable2[0])::]
        platdict[key] = value
    return platdict
    
#get each contents from configure file ---- map
def getSectionMap(strtmp,lable1 = partLable):
    tmp = strtmp.split(lable1[0])
    tmp = [elem for elem in tmp if len(elem) > 1]
    tmp = [elem for elem in tmp if elem.rfind(lable1[1]) > 0]
    sectionDict = {}
    for elem in tmp:
        key = elem[0:elem.find(lable1[1]):]
        value = elem[elem.find(endlineLable)+len(endlineLable)::]
        sectionDict[key] = value
    return sectionDict
    
#get detail options
def getValueMap(strtmp):
    tmp = strtmp.split(endlineLable)
    value = [elem for elem in tmp if len(elem) > 1]
    #print 'value ======%s' % value
    return value
    
#parse configure
def parseCFGByUserDefined (cfgFile, tag, option):
    f = open(cfgFile,"rb")
    strFileContent = f.read()
    f.close()
    vardict = {}
    var1 = getSectionMap(strFileContent)
    for k,v in var1.items():
        vardict[k] = v
    return vardict[tag]

#parse configure
def getPlatformAndSection (cfgFile):
    f = open(cfgFile,"rb")
    strFileContent = f.read()
    f.close()
    var1 = getPlatformMap(strFileContent)
    kvList = []
    for k in var1.items():
        kvList.append(k[0])
    return kvList

###########################################
##fileName: PM files 
##
###########################################
def parseXML(fileName):
    doc = ET.ElementTree(file=fileName)
    root = doc.getroot()
    all_node = root.getchildren()
    measData = all_node[1]
    children = measData.getchildren()
    measInfos = children[1:]
    return measInfos
  
###########################################
##tag: tag in cfg 
##option: which belong to tag
###########################################
def parseCFG(cfgFile, tag, option):
    config = ConfigParser.ConfigParser()
    config.readfp(open(cfgFile,"rb"))
    optionStr = config.get(tag, option)
    print 'optionStr = %s' % optionStr
    options = optionStr.split(';')
    return options
 
###########################################
##excelName: 
##statFields: 
##FileCount:
##dataDict:
###########################################
def writeData (excelName, statFields, FileCount, dataDict):  
    j = 2 
    for item in statFields:
        if dataDict.has_key(item):
            if FileCount == 1:
                excelName.write(0,j,item)
                excelName.write(FileCount,j,dataDict.get(item))
            else:
                excelName.write(FileCount,j,dataDict.get(item))
        j += 1 
  
def writeExcel(xlsFileName):
    
    global fileNameList
    
    #create excel work sheet
    book=xlwt.Workbook()
    
    kv = getPlatformAndSection(CFGPATH)

    print ('kv in cfg file is %s' % kv)
    for sheetname in kv:
        print('sheet name is %s' % sheetname)
        tagAndOption = sheetname.split('-')
        fields = parseCFGByUserDefined(CFGPATH, tagAndOption[0], 0)
        print ('fields is %s' % fields)
        if len(sheetname) > 31:
            sheetname = sheetname[0:30]
        #if sheetname=='AaaInterface-statisticsPerServer' or sheetname=='Gtp-statisticsPerPgw' or sheetname=='Gtp-QCI':
        #    pass
        #else:
        sheet = book.add_sheet(sheetname)
        sheet.write(0,0,"Group ID")
        sheet.write(0,1,"Time")
    
        i = 1
        z = 1
        #print ('collecting the %s data' % sheetname)
        #print 'fileNameList = %s' % fileNameList
        for fileName in fileNameList:
            print 'Collecting the %s data from file --%s' % (sheetname, (FILEPATH + fileName))
            #Group ID
            groupID = "G1"
        
            #Date
            #nameItem = fileName.split('-')
            nameItem = fileName.split('_')
            tt = nameItem[0].split('.')
            dateStr = tt[0][1:]
            yy = dateStr[:4]
            mm = dateStr[4:6]
            dd = dateStr[6:8]
            hh = tt[1][10:12]
            MM = tt[1][12:14]
            ss = '00'
            if hh == '00' and MM == '00':
                day = datetime.datetime.strptime(yy+'-'+mm+'-'+dd+' '+hh+':'+MM+':'+ss,'%Y-%m-%d %H:%M:%S') + datetime.timedelta(days=1)
                date = day.strftime('%Y-%m-%d %H:%M:%S')
            else:
                date = yy+'-'+mm+'-'+dd+' '+hh+':'+MM+':'+ss
            sheet.write(i,0,groupID)
            sheet.write(i,1,date)
              
            aaaDict = {}
            aaaObjDict = {}
            copyAAA = []
            measInfos = parseXML(FILEPATH+fileName)

            j = 2
            tempdict = {}
            tempList = []
            for item in measInfos:
                measValueList = []
                measTypeList = []
                tempValueList = []
                dict = {}
                measInfoTmp = item.getchildren()[2:]

                for item1 in measInfoTmp:
                    if valueStr.match(str(item1)):
                        measValueList.append(item1.getchildren())
                        tempValueList.append(item1)
                    else:
                        measTypeList.append(item1)
        
                length = len(measValueList)
                if length > 1:
                    tempObjLdn = ''

                    for measObj in tempValueList:
                        measObjLdn = measObj.attrib
                        instance = measObjLdn.get('measObjLdn')
                        #print ('objldn is %s' % instance)
                        temp = instance.split(',')
                        tempObjLdn = temp[2][0:-2]+'-'+temp[1]
                        
                        if temp[0] == 'AaaInterface' and temp[1] == 'statisticsPerServer':
                            copyAAA = measTypeList[:]
                            value = measObj.getchildren()
                            aaaDict[temp[3]] = value
                            aaaObjDict = {tempObjLdn:aaaDict}
                else:
                    measObjLdn = tempValueList[0].attrib
                    instance = measObjLdn.get('measObjLdn')
                    print ('objldn is %s' % instance)
                    temp = instance.split(',')
                    moduleInfo = temp[2][0:-2]
                    for item2 in measValueList[0]:
                        for item3 in measTypeList:
                            if item2.attrib.values()[0] == item3.attrib.values()[0]:
                                dict[item3.text] = item2.text
                    tempdict = {moduleInfo:dict}
                    tempList.append(tempdict)

            for dictItem in tempList:
                print ("dictItem is %s" % dictItem)
                print ("000 is %s" % dictItem.keys()[0])
                if dictItem.keys()[0] == sheetname or (dictItem.keys()[0])[0:30] == sheetname:
                    for field in fields:
                        print '=======field = %s' % field
                        if i == 1:
                            sheet.write(0,j,field)
                            sheet.write(i,j,(dictItem.values()[0]).get(field))
                        else:
                            sheet.write(i,j,(dictItem.values()[0]).get(field))
                        j += 1

            #data process for AAA
            if len(aaaObjDict) != 0:
                if sheetname == (aaaObjDict.keys()[0])[0:30]: 
                    if i == 1:                                   
                        sheet.write(0,2,"instance")
                    aaaValueDict = {}
                    aaaItem = aaaDict.items()
                    for elem in aaaItem:           
                        temp = list(elem)
                        for aaavalue in temp[1]:
                            for aaaField in copyAAA:
                                if aaaField.attrib.values()[0] == aaavalue.attrib.values()[0]:
                                    aaaValueDict[aaaField.text] = aaavalue.text
                        x = 3
                        #the second column of the worksheet is "Instance"
                        sheet.write(z,0,groupID)
                        sheet.write(z,1,date)
                        sheet.write(z,2,temp[0])
                        
                        #get AAA data
                        for field in fields:
                            print '=======field = %s' % field
                            if aaaValueDict.has_key(field):
                                #only one time for add field name
                                if z == 1 and i == 1:
                                    sheet.write(0,x,field)
                                    sheet.write(z,x,aaaValueDict.get(field))
                                else:
                                    sheet.write(z,x,aaaValueDict.get(field))
                            x += 1
                        z += 1
            i += 1
    book.save(xlsFileName) 
   
def mkdirs (path):
    #filter the first blank
    path = path.strip()
    #filter the last slash '/'
    #path = path.rstrip('/')
    #path = path+time.strftime('%Y-%m-%d')
    isExists = os.path.exists(path)
    if not isExists:
        os.makedirs(path)
        return path
    else:
        #directory aready exist
        return path 
    
if __name__ == '__main__':

    print ('Begin to collect data for epdg, please waiting...')
    for root, dirs, files in os.walk(FILEPATH):
        for fn in files:
            #print 'files are: %s' % fn
            fileNameList.append(fn)
    fileNameList.sort()

    date = time.strftime("%Y%m%d")
    hhmm = time.strftime("%H%M")
    #timezone = time.timezone 
    timezone = (os.popen("date -R | awk -F ' ' '{print $6}'")).readline().strip()
    hostname = (os.popen("hostname")).readline().strip()
    #print 'timezone: %s' % timezone
    #<yyyymmdd>.<hhmm><timezone>_<hostname>_epdg.xls
    mkdirs(outputPath)
    excelName = date + '.' + hhmm + str(timezone) + '_' + hostname + '_seg.xls'
    excelName = os.path.join(outputPath,excelName)
    writeExcel(excelName)

    print ('All data are collected successfully.')
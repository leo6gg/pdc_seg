#!/usr/bin/python
# -*- coding: utf-8 -*-
import os
import re
import xlwt
#import xlwt3 as xlwt
from xml.etree import ElementTree as ET
import ConfigParser
import time
import sys

curPath = os.getcwd()
tmpPath = os.path.abspath(os.path.join(curPath,os.path.pardir))
workPath = os.path.abspath(os.path.join(tmpPath,os.path.pardir))
configPath = curPath+'/config/'
#print configPath
globalFile = configPath + 'global.cfg'
#print ('globalFile == %s' % globalFile)
config = ConfigParser.ConfigParser()
config.readfp(open(globalFile,'rb'))
outputPath = curPath+ '/output/seg_data/'
FILEPATH = config.get("global","pmFilePath")
if FILEPATH == '':
    print('need to configure PM file path in '+ configPath + 'global.cfg')
    sys.exit()

#FILEPATH = '/home/mjw/Files/cp04/'
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
    #print ("========tmp in getPlatformMap  %s" % tmp )
    tmp = [elem for elem in tmp if len(elem) > 1]
    tmp = [elem for elem in tmp if elem.rfind(lable1[1]) > 0]
    platdict = {}
    #print ("========tmp in getPlatformMap  %s" % tmp )
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
    dict3 = {}
    for k,v in var1.items():
        vardict[k] = v
    return vardict[tag]


def getPlatformAndSection (cfgFile):
    f = open(cfgFile,"rb")
    strFileContent = f.read()
    f.close()
    vardict = {}
    var1 = getPlatformMap(strFileContent)
    kvList = []
	
    for k,v in var1.items():
        #print('=======module key===%s' % k)
        kvList.append(k)
    return kvList

###########################################
##fileName: PM files 
##
###########################################
def parseXML(fileName):
    measTypeList = []
    measValueList = []
    dict = {}
    valueStr = re.compile('.+measValue.+')
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
    options = optionStr.split(';')
    return options
  
  
def writeExcel(xlsFileName):
    
    global fileNameList

    #create excel work sheet
    book=xlwt.Workbook()
    
    kv = getPlatformAndSection(CFGPATH)
    print('===========kv is %s' % kv)
	
    for sheetname in kv:
        print('===========sheet name is %s' % sheetname)
        tagAndOption = sheetname.split('-')
        print('===========tag name is %s' % tagAndOption[0])

        fields = parseCFGByUserDefined(CFGPATH, tagAndOption[0], 0)

        if len(sheetname) > 31:
            sheetname = sheetname[0:30]

        sheet = book.add_sheet(sheetname)
        sheet.write(0,0,"Group ID")
        sheet.write(0,1,"Time")
    
        i = 1
        z = 1
        y = 1
        
        #print 'fileNameList = %s' % fileNameList
        for fileName in fileNameList:
            print '==============Collect the %sth file --%s' % (i, (FILEPATH + fileName))
            #Group ID
            groupID = "G00"
        
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
            #print('11111111111111111111111111111111111111 i = %s' % i)
            if sheetname=='AaaInterface-statisticsPerServ' or sheetname=='Gtp-statisticsPerPgw' or sheetname=='Gtp-QCI':
                pass
            else:
                sheet.write(i,0,groupID)
                sheet.write(i,1,date)
              
            aaaDict = {}
            aaaObjDict = {}
            
            measInfos = parseXML(FILEPATH+fileName)

            j = 2
            tempdict = {}
            tempList = []
            for item in measInfos:
                #print '===========item in measInfos ==%s' % item
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
                #print '===========measValueList==%s' % measValueList

                length = len(measValueList)
                print ('=========length is %s=======' % length)

                if length > 1:
                    tempObjLdn = ''
                    copyall = []
                    allDict = {}
                    for measObj in tempValueList:
                        measObjLdn = measObj.attrib
                        instance = measObjLdn.get('measObjLdn')
                        temp = instance.split(',')
                        tempObjLdn = temp[0]+'-'+temp[1]
                        
                        if temp[0] == 'AaaInterface' and temp[1] == 'statisticsPerServer':
                            copyAAA = measTypeList[:]
                            value = measObj.getchildren()
                            aaaDict[temp[3]] = value
                            aaaObjDict = {tempObjLdn:aaaDict}
                else:
                    measObjLdn = tempValueList[0].attrib
                    instance = measObjLdn.get('measObjLdn')
                    temp = instance.split(',')
                    moduleInfo = temp[0]+'-'+temp[1]
                    for item2 in measValueList[0]:
                        for item3 in measTypeList:
                            if item2.attrib.values()[0] == item3.attrib.values()[0]:
                                dict[item3.text] = item2.text
                    tempdict = {moduleInfo:dict}
                    tempList.append(tempdict)

            for dictItem in tempList:
                print ('=======dictItem.keys()[0]=======item=========%s,%s' % (dictItem.keys()[0],sheetname))
                if dictItem.keys()[0] == sheetname or (dictItem.keys()[0])[0:30] == sheetname:
                    for field in fields:
                        if i == 1:
                            sheet.write(0,j,field)
                            sheet.write(i,j,(dictItem.values()[0]).get(field))
                        else:
                            sheet.write(i,j,(dictItem.values()[0]).get(field))
                        j += 1
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
            fileNameList.append(fn)

    fileNameList.sort()
    #get the newest file
    lastFile = fileNameList[-1]
    dateInfo = lastFile.split('_')
    tmp = dateInfo[0].split('.')
    ymd = tmp[0][1:]
    HM = tmp[1][10:14]
    filetimezone = tmp[1][4:9]
    dateStr = ymd + HM
    newestFileList = []
    for item in fileNameList:
        fileNameStr = item.split('_')
        yymmdd = (fileNameStr[0].split('.')[0])[1:]
        fileDate = yymmdd + ((fileNameStr[0].split('.')[1]))[10:14]
        if str(fileDate) == str(dateStr):
            newestFileList.append(item)
    
    fileNameList = newestFileList[:]
    print ('fileNameList %s ' % fileNameList)
    
    date = time.strftime("%Y%m%d")
    hhmm = time.strftime("%H%M")

    timezone = (os.popen("date -R | awk -F ' ' '{print $6}'")).readline().strip()
    hostname = (os.popen("hostname")).readline().strip()

    mkdirs(outputPath)
    excelName = ymd + '.' + HM + filetimezone + '_' + hostname + '_seg_latest.xls'
    excelName = os.path.join(outputPath,excelName)
    writeExcel(excelName)
    print ('Collect data successfully.')

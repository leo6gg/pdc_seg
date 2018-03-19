#!/usr/bin/python
# -*- coding: utf-8 -*-
import os
import re
import xlwt
#import xlwt3 as xlwt
from xml.etree import ElementTree as ET
import ConfigParser
import time
import xlsxwriter
import xlrd
import string
import datetime
import sys
'''
curPath = os.getcwd()
tmpPath = os.path.abspath(os.path.join(curPath,os.path.pardir))
workPath = os.path.abspath(os.path.join(tmpPath,os.path.pardir))
configPath = curPath+'/config/'
#print configPath
globalFile = configPath + 'global.cfg'
#print ('globalFile == %s' % globalFile)
config = ConfigParser.ConfigParser()
config.readfp(open(globalFile,'rb'))
outputPath = curPath + '/output/cp0405_data/'
FILEPATH = config.get("global","pmFilePath")
if FILEPATH == '':
    print('need to configure PM file path in '+ configPath + 'global.cfg')
    sys.exit()
'''
config = ConfigParser.ConfigParser()
config.readfp(open('/home/mjw/pdctool/config/global.cfg','rb'))
#test = config.get("global","tempInfo")
#print 'test == %s' % test
#config.set('global','tempInfo','ddsssssss')

FILEPATH = '/home/mjw/Files/cp04/pm/'
#FILEPATH = '/home/mjw/Files/15b/'
#CFGPATH = configPath + 'chart_config.cfg'
CFGPATH = '/home/mjw/pdctool/config/chart_config.cfg'
fileNameList = []
daylastfile = []


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
def getSectionMap(strtmp,lable1 = sectionLable):
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
    var1 = getPlatformMap(strFileContent)
    for k,v in var1.items():
        var2 = getSectionMap(v)
        dict3 = {}
        for k2,v2 in var2.items():
            var3 = getValueMap(v2)
            dict3[k2] = var3
        vardict[k] = dict3
    return vardict[tag][option]

def makeListElementUnique (L):
    newList = []
    for item in L:
        if item not in newList:
            newList.append(item)
    return newList

#get groups
def getGroups(fileNameList):
    groups = []
    for fileName in fileNameList:
        #Group ID
        fileStr = fileName.split('_')
        group = fileStr[1].split('-')[-1]
        groupID = "G"+group
        groups.append(groupID)
        
    group = makeListElementUnique(groups)
    return group
                  

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
    print 'optionStr = %s' % optionStr
    options = optionStr.split(';')
    return options

#parse configure
def getPlatformAndSection (cfgFile):
    f = open(cfgFile,"rb")
    strFileContent = f.read()
    f.close()
    vardict = {}
    var1 = getPlatformMap(strFileContent)
    kvList = []
    for k,v in var1.items():
        #print('=======module key===%s' % k)
        var2 = getSectionMap(v)
        for k2,v2 in var2.items():
            #print('++++++child module key===%s' % k2)
            kvList.append(k + '-' + k2)
    return kvList
 
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

#get aaa and gtp instance count
def getGtpAndAaaInstance (fileName):    
    aaaInstance = []
    gtpInstance = []
    qciInstance = []
    tempValueList = []
    measValueList = []
    measInfos = parseXML(FILEPATH+fileName)
    for item in measInfos:
        measInfoTmp = item.getchildren()[2:]
        for item1 in measInfoTmp:
            if valueStr.match(str(item1)):
                measValueList.append(item1.getchildren())
                tempValueList.append(item1)
            
        length = len(measValueList)
        if length > 1:
            for measObj in tempValueList:
                measObjLdn = measObj.attrib
                instance = measObjLdn.get('measObjLdn')
                temp = instance.split(',')
                #print 'temp = %s' % temp
                if temp[0] == 'AaaInterface' and len(temp) == 4:
                    aaaInstance.append(temp[3])
                if temp[0] == 'Gtp' and len(temp) == 4 and temp[1] == 'statisticsPerPgw':
                    gtpInstance.append(temp[3])
                if temp[0] == 'Gtp' and len(temp) == 4 and temp[1] == 'QCI':
                    qciInstance.append(temp[3])
        else:
            pass
    return aaaInstance, gtpInstance, qciInstance

##################
  
def writeExcel(xlsFileName):
    
    global fileNameList

    #create excel work sheet
    book=xlwt.Workbook()
    
    groups = getGroups(fileNameList)
    print 'groups=====%s' % groups
    kv = getPlatformAndSection(CFGPATH)
    for groupItem in groups:
        for sheetname in kv:
            #print('sheet name is %s' % sheetname)
            tagAndOption = sheetname.split('-')
            fields = parseCFGByUserDefined(CFGPATH, tagAndOption[0], tagAndOption[1])
            #print('@@@@@@@@@@fields are %s' % fields)
            if len(sheetname) > 31:
                sheetname = sheetname[0:30]
            if 'SystemManagement' in sheetname:
                sheetname = sheetname.replace('SystemManagement','SysMgmt')     	
            sheet = book.add_sheet(groupItem +'_'+ sheetname, cell_overwrite_ok=True)
            sheet.write(0,0,"Group ID")
            sheet.write(0,1,"Time")
        
            i = 1
            z = 1
            y = 1
            print ('collecting the %s %s data' % (groupItem,sheetname))
            #print 'fileNameList = %s' % fileNameList
            for fileName in fileNameList:
                
                #print 'Collect the %sth file --%s' % (i, (FILEPATH + fileName))
                #Group ID
                fileStr = fileName.split('_')
                group = fileStr[1].split('-')[-1]
                groupID = "G"+group
            
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
                if groupID != groupItem:
                    continue
                else:
                    aaaDict = {}
                    aaaObjDict = {}
                    gtpDict = {}
                    gtpObjDict = {}
                    gtpQciDict = {}
                    gtpObgQciDict = {}
                    copyGtpQCI = []
                    copyGTP = []
                    copyAAA = []
                    measInfos = parseXML(FILEPATH+fileName)
                    #print '===========measTypeList==%s' % measTypeList
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
                        #print '===========measValueList==%s' % measValueList
                        
                       
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
                            elif temp[0] == 'Gtp' and temp[1] == 'statisticsPerPgw':
                                copyGTP = measTypeList[:]
                                value = measObj.getchildren()
                                gtpDict[temp[3]] = value
                                gtpObjDict = {tempObjLdn:gtpDict}
                            elif temp[0] == 'Gtp' and temp[1] == 'QCI':
                                copyGtpQCI = measTypeList[:]
                                value = measObj.getchildren()
                                gtpQciDict[temp[3]] = value
                                gtpObgQciDict = {tempObjLdn : gtpQciDict}
                            
                            else:
                                for item2 in measValueList[0]:
                                    for item3 in measTypeList:
                                        if item2.attrib.values()[0] == item3.attrib.values()[0]:
                                            dict[item3.text] = item2.text
                                tempdict = {tempObjLdn:dict}
                                tempList.append(tempdict)
                            
                        
                        #print ('+++++++++++++dict = %s' % dict)
                        #print ('++++++@@@@@+++++++tempdict = %s' % tempdict)
                    #print ('++++++#####+++++++tempList = %s' % tempList)
                    #print ('++++++#####+++++++gtpQciDict = %s' % gtpQciDict)
                    a = 0                      
                    for dictItem in tempList:
                        #print ('=======dictItem.keys()[0]=======item=========%s,%s' % (dictItem.keys()[0],sheetname))
                        if 'SysMgmt' in sheetname:
                            sheetname = 'SystemManagement-General'
   
                        #if dictItem.keys()[0] == sheetname or (dictItem.keys()[0])[0:30] == sheetname:
                        if dictItem.keys()[0] == sheetname:
                            record = []
                            for field in fields:
                                ##如果是 CC要取差值 (如果 15分是 a1，30分变a2,那30分显示的值是 a2-a1)
                                if 'CC' in field:
                                    print ('aaaa--field = %s' % field)
                                    field = field.split('--')[0]
                                    print ('aaaaaaaaaaaaaaa==') 
                                    if i == 1:
                                        sheet.write(0,j,field)
                                        #sheet.write(i,j,(dictItem.values()[0]).get(field))
                                        sheet.write(i,j,0)
                                        sheet.write(i,0,groupID)
                                        sheet.write(i,1,date)
                                        record.append((dictItem.values()[0]).get(field))
                                    else:
                                        #print ('aaaa--record = %s' % record)
                                        result = config.get('record','lastRecord')
                                        #print ('aaaa--result = %s' % result)
                                        temp = int((dictItem.values()[0]).get(field)) - int(result[a])
                                        sheet.write(i,j,temp)
                                        sheet.write(i,0,groupID)
                                        sheet.write(i,1,date)
                                        record.append((dictItem.values()[0]).get(field))
                                    
                                    
                                    #print ('iiiiiiiiiiiii==%s' % i)
                                    a += 1
                                else:
                                    if i == 1:
                                        sheet.write(0,j,field)
                                        sheet.write(i,j,(dictItem.values()[0]).get(field))
                                        sheet.write(i,0,groupID)
                                        sheet.write(i,1,date)
                                    else:
                                        sheet.write(i,j,(dictItem.values()[0]).get(field))
                                        sheet.write(i,0,groupID)
                                        sheet.write(i,1,date)
                                
                                j += 1
                            
                        if len(record) != 0:
                            config.set('record','lastRecord',record)
                            config.write(open('/home/mjw/pdctool/config/global.cfg','w'))      
                            #print 'record = %s' % record
                           
                    #print 'aaaDict = %s' % aaaDict
                    #print 'gtpDict = %s' % gtpDict 
                    
                    
                    
                    #data process for AAA
                    if len(aaaObjDict) != 0:
                        if 'AaaInterface' in sheetname:
                            sheetname = 'AaaInterface-statisticsPerServ'
                        if sheetname == (aaaObjDict.keys()[0])[0:30]: 
                            #print ('aaaObjDict.keys()[0])[0:30]==========%s' % (aaaObjDict.keys()[0])[0:30])
                            if i == 1:                                   
                                sheet.write(0,2,"Instance")
                            aaaValueDict = {}
                            aaaItem = aaaDict.items()
                            #print '========aaaItem= %s' % aaaItem
                            record = []
                            b = 0
                            for elem in aaaItem:           
                                temp = list(elem)
                                for aaavalue in temp[1]:
                                    for aaaField in copyAAA:
                                        if aaaField.attrib.values()[0] == aaavalue.attrib.values()[0]:
                                            aaaValueDict[aaaField.text] = aaavalue.text
                                print '=======aaaValueDict = %s' %  aaaValueDict           
                                x = 3
                                #the second column of the worksheet is "Instance"
                                #print '=======groupID = %s' %  groupID
                                sheet.write(z,0,groupID)
                                sheet.write(z,1,date)
                                sheet.write(z,2,temp[0])
                                
                                #get AAA data
                                
                                for field in fields:
                                    #print '=======field = %s' % field
                                    
                                    if 'CC' in field:
                                        print ('cc field === %s' % field)
                                        field = field.split('--')[0]
                                        record.append(aaaValueDict.get(field))
                                        if i ==1:
                                            if aaaValueDict.has_key(field):
                                                #only one time for add field name
                                                if z == 1 and i == 1:
                                                    sheet.write(0,x,field)
                                                    sheet.write(z,x,0)
                                                else:
                                                    sheet.write(z,x,0)
                                            x += 1
                                        else:
                                            result = config.get('record','lastRecord')
                                            print ('aaaa--result = %s' % result)
                                            print ('bbbb--bbbbb = %s' % b)
                                            if aaaValueDict.has_key(field):
                                                #only one time for add field name
                                                temp = int(aaaValueDict.get(field)) - int(result[b])
                                                if z == 1 and i == 1:
                                                    sheet.write(0,x,field)
                                                    sheet.write(z,x,temp)
                                                else:
                                                    sheet.write(z,x,temp)
                                            x += 1
                                        b += 1    
                                    else:
                                        if aaaValueDict.has_key(field):
                                            #only one time for add field name
                                            if z == 1 and i == 1:
                                                sheet.write(0,x,field)
                                                sheet.write(z,x,aaaValueDict.get(field))
                                            else:
                                                sheet.write(z,x,aaaValueDict.get(field))
                                        x += 1
                                z += 1
                            print 'record recordrecordrecordrecord= %s' % record
                            if len(record) != 0:
                                print ('@@@@@@@@@@@@@@@@@@@@@@%s'% record)
                                config.set('record','lastRecord','record')
                                config.write(open('/home/mjw/pdctool/config/global.cfg','w')) 
                    #data process for Gtp and QCI
                    if 'Gtp-Pgw' in sheetname:
                        #print ('sheetname==========%s' % sheetname)
                        sheetname = 'Gtp-statisticsPerPgw'
                    if sheetname == gtpObjDict.keys()[0]:
                        #print ('aaaObjDict.keys()[0])[0:30]==========%s' % (gtpObjDict.keys()[0]))
                        if i == 1:                                   
                            sheet.write(0,2,"instance")
                            
                        gtpValueDict = {}
                        gtpItem = (gtpObjDict.values()[0]).items()
                        #print '========gtpItem= %s' % gtpItem
                        for elem in gtpItem:
                            temp = list(elem)
                            for gtpvalue in temp[1]:
                                for gtpField in copyGTP:
                                    if gtpField.attrib.values()[0] == gtpvalue.attrib.values()[0]:
                                        gtpValueDict[gtpField.text] = gtpvalue.text
                            #print '=======gtpValueDict = %s' %  gtpValueDict           
                            x = 3
                            #the second column of the worksheet is "Instance"
                            sheet.write(y,0,groupID)
                            sheet.write(y,1,date)
                            sheet.write(y,2,temp[0])
                            #get Gtp data
                            for field in fields:
                                #print '=======field = %s' % field
                                if gtpValueDict.has_key(field):
                                    #only one time for add field name
                                    if y == 1 and i == 1:
                                        sheet.write(0,x,field)
                                        sheet.write(y,x,gtpValueDict.get(field))
                                    else:
                                        sheet.write(y,x,gtpValueDict.get(field))
                                x += 1
                            y += 1
                            
                    if sheetname == gtpObgQciDict.keys()[0]:
                        if i == 1:                                   
                            sheet.write(0,2,"instance")
                            
                        gtpQciValueDict = {}
                        gtpItem = (gtpObgQciDict.values()[0]).items()
                        #print '========gtpItem= %s' % gtpItem
                        for elem in gtpItem:
                            temp = list(elem)
                            for gtpvalue in temp[1]:
                                for gtpField in copyGtpQCI:
                                    if gtpField.attrib.values()[0] == gtpvalue.attrib.values()[0]:
                                        gtpQciValueDict[gtpField.text] = gtpvalue.text
                            #print '==****************=====gtpQciValueDict = %s' %  gtpQciValueDict.keys()           
                            x = 3
                            #the second column of the worksheet is "Instance"
                            sheet.write(y,0,groupID)
                            sheet.write(y,1,date)
                            sheet.write(y,2,temp[0])
                            #get Gtp data
                            for field in fields:
                                #print '=======field = %s' % field
                                if gtpQciValueDict.has_key(field):
                                    #only one time for add field name
                                    if y == 1 and i == 1:
                                        sheet.write(0,x,field)
                                        sheet.write(y,x,gtpQciValueDict.get(field))
                                    else:
                                        sheet.write(y,x,gtpQciValueDict.get(field))
                                x += 1
                            y += 1
                    
                          
                    i += 1  
        book.save(xlsFileName) 

#############################################################
def chart_series(sheetName, row, cur_row):
    chart.add_series({
                      'categories': [sheetName,1,1,row,1],
                      'values':     [sheetName,1,cur_row,row,cur_row],
                      'name': [sheetName,0,cur_row],
                     })
    
def chart_series_aaagtp(sheetName, row, cur_row):
    chart.add_series({
                      'categories': [sheetName,1,1,row,1],
                      'values':     [sheetName,1,cur_row,row,cur_row],
                      'name': [sheetName,0,cur_row],
                     })

def sumdata (sheetFrom, sheetTo, cur_row, cur_col, worksheet):
    charList = ['A','B','C','D','E','F','G','H','I','J','K','L','M','N','O','P','Q','R','S','T','U','V','W','X','Y','Z']
    for i in range(cur_col):
        worksheet.write_formula(charList[i]+cur_row, \
     '=SUM('+sheetFrom+':'+sheetTo+'!'+charList[i+1]+cur_row+':'+charList[i+1]+cur_row+')')    

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
    fileDateList = []
    for item in fileNameList:
        dateStrList = item.split('.')
        dateStr = dateStrList[0][1:]
        year = dateStr[0:4]
        mon = dateStr[4:6]
        day = dateStr[6:8]
        tempDate = year+'-'+mon+'-'+day
        fileDateList.append(tempDate)
    fileDateList = makeListElementUnique(fileDateList)
    
    '''
    for temp in fileDateList:
        dayfile = []
        for item in fileNameList:
            dateStrList = item.split('.')
            dateStr = dateStrList[0][1:]
            year = dateStr[0:4]
            mon = dateStr[4:6]
            day = dateStr[6:8]
            tempDate = year+'-'+mon+'-'+day
            if temp == tempDate:
                dayfile.append(item)
        daylastfile.append(dayfile[-1])
                    
    print('daylastfile ===== %s ' % daylastfile)    
    '''    
        
    date = time.strftime("%Y%m%d")
    hhmm = time.strftime("%H%M")
    #timezone = time.timezone 
    timezone = (os.popen("date -R | awk -F ' ' '{print $6}'")).readline().strip()
    hostname = (os.popen("hostname")).readline().strip()
    #print 'timezone: %s' % timezone
    #<yyyymmdd>.<hhmm><timezone>_<hostname>_epdg.xls
    #mkdirs(outputPath)
    excelName = date + '.' + hhmm + str(timezone) + '_' + hostname + '_wmg.xls'
    #excelName = os.path.join(outputPath,excelName)
    writeExcel(excelName)
    sys.exit()    
    print ('data analyzing......')
    #get aaa and gtp instance count
    aaaInstance, gtpInstance, qciInstance = getGtpAndAaaInstance(fileNameList[0])
    aaaInstance = makeListElementUnique(aaaInstance)
    gtpInstance = makeListElementUnique(gtpInstance) 
    qciInstance = makeListElementUnique(qciInstance)
    print ('aaaInstance === %s' % aaaInstance)
    print ('gtpInstance === %s' % gtpInstance) 
    print ('qciInstance === %s' % qciInstance)
    
    workbook = xlsxwriter.Workbook(os.path.join(outputPath,date+'-'+hhmm+timezone+'_'+hostname+'_wmg_graph.xlsx'))
    
    #data = xlrd.open_workbook('20150513.1659+0800_ubuntu_wmg.xls')
    data = xlrd.open_workbook(excelName)
    sheets = data.sheet_names()
    date_format = workbook.add_format({'num_format': 'dd/mm/yyyy hh:mm'})
    	
    ##############################################
    #deal with Sc_Call_Distribution, it only need the last data of every day
    allDistribution = workbook.add_worksheet('Sc-Call_Distribution')
    DistributionChartPosition = 1
    DistributionFlag = 0
    for item in sheets:        
        if str(item)[3:] == 'Sc-Call_Distribution':
            print ('Sc-Call_Distribution ===== %s' % item)
            disSheetName = 'day'+'_'+str(item)
            worksheet = workbook.add_worksheet(disSheetName)
            chart = workbook.add_chart({'type': 'line'})
            table = data.sheet_by_name(item)
            worksheet.write_row(0,0,table.row_values(0))
            nrows = table.nrows
            ncols = table.ncols
            lastDateList = []
            for tem in fileDateList:
                tempList = []
                for i in range(1,nrows):
                    cell_value = table.cell_value(i,1)
                    cell_date = cell_value.split(' ')                    
                    if tem == cell_date[0]:
                        tempList.append(table.cell_value(i,1))
                    
                      
                #tempList.sort()
                print ('tempList ===== %s' % tempList)
                if len(tempList) != 0 :
                    lastDateList.append(tempList[-1])
            print ('lastDateList ===== %s' % lastDateList)
            c = 1
            for lastDate in lastDateList:
                for i in range(1,nrows):
                    cell_value = table.cell_value(i,1)                    
                    if lastDate == cell_value:
                        worksheet.write_row(c,0,table.row_values(i))
                        c += 1
            
            for cur_row in range(2,ncols):
                    chart_series(disSheetName, c, cur_row)
        
            chart.set_size({'width': 500, 'height': 287})
            chart.set_title ({'name': item+'_stat.'})
            #chart.set_y_axis({'name': 'count'})
            chart.set_y_axis({
                'name': 'Units',
                'name_font': {
                    'name': 'Century',
                    'color': 'red'
                },
                'num_font': {
                    'bold': True,
                    'italic': True,
                    'underline': True,
                    'color': '#7030A0',
                },
            })
          
            if DistributionFlag % 2 == 0:
                allDistribution.insert_chart(DistributionChartPosition, 0, chart) 
                DistributionChartPosition += 16
            else: 
                DistributionChartPosition -= 16
                allDistribution.insert_chart(DistributionChartPosition, 9, chart)
                DistributionChartPosition += 16
            DistributionFlag += 1            
                
    #deal with Sc_Call_Distribution end
    ##############################################   
    
    aaaSheets = []
    gtpSheets = []
    qciSheets = []
    aaastr = re.compile('.+AaaInterface.+')
    gtpstr = re.compile('.+Gtp-Pgw.+')
    qcistr = re.compile('.+Gtp-QCI')
    sheetNamePart = ['SysMgmt-cpuAndMem','SysMgmt-throughput_M','SysMgmt-throughput_K','Sc-General','Sc-Emergency_Call',\
        'Sc-Call_Duration','Ipsec-General','Ipsec-Summary','Dns-General','Gtp-General','Gtp-Summary']
    print ('sheet sheets is %s' % sheets)
    for namePart in sheetNamePart:
        allworksheet = workbook.add_worksheet(namePart)
        chartPosition = 1
        flag = 0
        for sheet in sheets:
            if aaastr.match(str(sheet)):
                aaaSheets.append(sheet)
            elif gtpstr.match(str(sheet)):
                gtpSheets.append(sheet)
            elif qcistr.match(str(sheet)):
                qciSheets.append(sheet)
            elif namePart in sheet:
                print ('sheet name is %s' % sheet)
                chart = workbook.add_chart({'type': 'line'})
                worksheet = workbook.add_worksheet(sheet)
                table = data.sheet_by_name(sheet)
                nrows = table.nrows
                ncols = table.ncols
                #print ('sheet ncols is %s' % ncols)
                for i in xrange(nrows):
                    for j in xrange(ncols):
                        cell_value = table.cell_value(i,j,)
                        #type(eval('33.33')) == float
                        if type(cell_value) == unicode and cell_value.isdigit():
                            worksheet.write(i,j,string.atoi(cell_value))
                        elif i >= 1 and j == 1:
                            worksheet.write(i,j,cell_value,date_format)
                        else:
                            worksheet.write(i,j,cell_value)
        
                #print ('nrows is %s' % nrows)            
                for cur_row in range(2,ncols):
                    chart_series(sheet, nrows, cur_row)
        
                chart.set_size({'width': 500, 'height': 287})
                chart.set_title ({'name': sheet+'_stat.'})
                #chart.set_y_axis({'name': 'count'})
                chart.set_y_axis({
                    'name': 'Units',
                    'name_font': {
                        'name': 'Century',
                        'color': 'red'
                    },
                    'num_font': {
                        'bold': True,
                        'italic': True,
                        'underline': True,
                        'color': '#7030A0',
                    },
                })
                '''
                chart.set_x_axis({
                             'date_axis': True,
                             'min': minDate,
                             'max': maxDate,
                             })
                '''
                #worksheet.insert_chart('G4', chart)
                #print ('flag ==== %s ' % flag)
                if flag % 2 == 0:
                    allworksheet.insert_chart(chartPosition, 0, chart) 
                    chartPosition += 16
                else: 
                    chartPosition -= 16
                    allworksheet.insert_chart(chartPosition, 9, chart)
                    chartPosition += 16
                flag += 1
                
                
    
    
    aaaSheets =  makeListElementUnique(aaaSheets) 
    gtpSheets =  makeListElementUnique(gtpSheets) 
    qciSheets =  makeListElementUnique(qciSheets)
    print ('aaaSheet ==== %s' % aaaSheets)
    print ('gtpSheet ==== %s' % gtpSheets) 
    print ('qciSheets ==== %s' % qciSheets)
    #sum_sheet_wb = workbook.add_worksheet('sum_sheet')
    tmpSheetList = []
    
    for aaaSheet in aaaSheets:
        preName = aaaSheet[0:2]
        sheetnamestr = aaaSheet.split('-')
        sheet_for_aaa = workbook.add_worksheet('g'+aaaSheet[1:2]+'_'+sheetnamestr[1])
        table = data.sheet_by_name(aaaSheet)
        nrows = table.nrows
        ncols = table.ncols
        print ('sheet ncols is %s' % ncols) 
        print ('sheet nrows is %s' % nrows)
        chartPosition = 1
        for instance in aaaInstance:
            partInstance = instance.split('=')
            copySheetName = aaaSheet[0:2]+'_'+sheetnamestr[1]+'_'+partInstance[1]
            worksheet = workbook.add_worksheet(copySheetName)
            worksheet.write_row(0,0,table.row_values(0))
            chart = workbook.add_chart({'type': 'line'})
            i = 1
            for rownum in range(1, nrows):
                if instance == table.cell(rownum,2).value:
                    print ('i === %s' % i)
                    for colnum in xrange(ncols):
                        cell_value = table.cell_value(rownum,colnum,)
                        if type(cell_value) == unicode and cell_value.isdigit():
                            worksheet.write(i, colnum, string.atoi(cell_value))
                        else:
                            worksheet.write(i,colnum,cell_value)

                    i += 1
 
            #tmpSheetList.append(aaaSheet+instance)
            tmpSheetList.append(worksheet)
            #print ('tmpSheetList is %s' % tmpSheetList)            
            for cur_row in range(3,ncols):
                chart_series_aaagtp(copySheetName, nrows/len(aaaInstance), cur_row)
                
            chart.set_size({'width': 500, 'height': 287})
            chart.set_title ({'name': aaaSheet[3:]+'_'+instance})
            #chart.set_y_axis({'name': 'count'})
            chart.set_y_axis({
                'name': 'Units',
                'name_font': {
                    'name': 'Century',
                    'color': 'red'
                },
                'num_font': {
                    'bold': True,
                    'italic': True,
                    'underline': True,
                    'color': '#7030A0',
                },
            })
            
            #chart.set_x_axis({
            #             'date_axis': True,
            #             'min': minDate,
            #             'max': maxDate,
            #             })
            
            #worksheet.insert_chart('G4', chart)
            sheet_for_aaa.insert_chart(chartPosition, 0, chart)
            chartPosition += 16
    
    #os.exit()
    ################################################################################       
    #process for GTP
    for gtpSheet in gtpSheets:
        preName = gtpSheet[0:2]
        sheetnamestr = gtpSheet.split('-')
        sheet_for_gtp = workbook.add_worksheet('g'+gtpSheet[1:2]+'_'+sheetnamestr[1])
        table = data.sheet_by_name(gtpSheet)
        nrows = table.nrows
        ncols = table.ncols
        chartPosition = 1
        for instance in gtpInstance:
            partInstance = instance.split('=')
            if partInstance[1] == 'ffff:ffff:ffff:ffff:ffff:ffff:ffff:ffff':
                sheetname = preName+'_'+sheetnamestr[1]+'_'+'ffff.ffff.ffff.ffff.ffff.ffff.ffff.ffff'
            else:
                sheetname = preName+'_'+sheetnamestr[1]+'_'+instance
            if len(sheetname) > 31 :
                sheetname = sheetname[0:31]
            else:
                sheetname = sheetname
            worksheet = workbook.add_worksheet(sheetname)
            worksheet.write_row(0,0,table.row_values(0))
            chart = workbook.add_chart({'type': 'line'})
            i = 1
            for rownum in range(1, nrows):
                if instance == table.cell(rownum,2).value:
                    print ('gtp=== i === %s' % i)
                    for colnum in xrange(ncols):
                        cell_value = table.cell_value(rownum,colnum,)
                        if type(cell_value) == unicode and cell_value.isdigit():
                            worksheet.write(i, colnum, string.atoi(cell_value))
                        else:
                            worksheet.write(i,colnum,cell_value)

                    i += 1 
            #print ('nrows is %s' % nrows)            
            for cur_row in range(3,ncols):
                chart_series_aaagtp(sheetname, nrows/len(gtpInstance), cur_row)
                
            chart.set_size({'width': 500, 'height': 287})
            #chart.set_title ({'name': sheetname[2:]})
            chart.set_title ({'name': sheetnamestr[1]+'_'+instance})
            #chart.set_y_axis({'name': 'count'})
            chart.set_y_axis({
                'name': 'Units',
                'name_font': {
                    'name': 'Century',
                    'color': 'red'
                },
                'num_font': {
                    'bold': True,
                    'italic': True,
                    'underline': True,
                    'color': '#7030A0',
                },
            })
            
            #chart.set_x_axis({
            #             'date_axis': True,
            #             'min': minDate,
            #             'max': maxDate,
            #             })
            
            #worksheet.insert_chart('G4', chart)
            sheet_for_gtp.insert_chart(chartPosition, 0, chart)
            chartPosition += 16
    
    ################################################################################       
    #process for QCI
    for qciSheet in qciSheets:
        preName = qciSheet[0:2]
        sheetnamestr = qciSheet.split('-')
        sheet_for_qci = workbook.add_worksheet('g'+qciSheet[1:2]+'_'+sheetnamestr[1])
        table = data.sheet_by_name(qciSheet)
        nrows = table.nrows
        ncols = table.ncols
        #print ('sheet ncols is %s' % ncols) 
        #print ('sheet nrows is %s' % nrows)
        chartPosition = 1
        for instance in qciInstance:
            partInstance = instance.split('=')
            copySheetName = qciSheet[0:2]+'_'+sheetnamestr[1]+'_'+partInstance[1]
            worksheet = workbook.add_worksheet(copySheetName)
            worksheet.write_row(0,0,table.row_values(0))
            chart = workbook.add_chart({'type': 'line'})
            i = 1
            for rownum in range(1, nrows):
                if instance == table.cell(rownum,2).value:
                    for colnum in xrange(ncols):
                        cell_value = table.cell_value(rownum,colnum,)
                        if type(cell_value) == unicode and cell_value.isdigit():
                            worksheet.write(i, colnum, string.atoi(cell_value))
                        else:
                            worksheet.write(i,colnum,cell_value)

                    i += 1
            
            for cur_row in range(3,ncols):
                chart_series_aaagtp(copySheetName, nrows/len(qciInstance), cur_row)
                
            chart.set_size({'width': 500, 'height': 287})
            chart.set_title ({'name': qciSheet[3:]+'_'+instance})
            #chart.set_y_axis({'name': 'count'})
            chart.set_y_axis({
                'name': 'Units',
                'name_font': {
                    'name': 'Century',
                    'color': 'red'
                },
                'num_font': {
                    'bold': True,
                    'italic': True,
                    'underline': True,
                    'color': '#7030A0',
                },
            })
            
            #chart.set_x_axis({
            #             'date_axis': True,
            #             'min': minDate,
            #             'max': maxDate,
            #             })
            
            #worksheet.insert_chart('G4', chart)
            sheet_for_qci.insert_chart(chartPosition, 0, chart)
            chartPosition += 16
    
    print ('data charts are generating......')
    #hide the worksheets
    
    worksheets = workbook.worksheets()
    #print ('worksheets count = %s' % worksheets)
    
    for worksheet in worksheets:
        if 'G' in worksheet.get_name():
            worksheet.hide()
    
    workbook.close()
    
    print ('Collect data successfully.')

#!/usr/bin/python
# -*- coding: utf-8 -*-
import os, platform
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


curPath = os.getcwd()
curPath = '/opt/leo/pdctool20160104'

if platform.system() == 'Windows':
    curPath = '\\\\192.168.56.101\\Leo\\pdctool20160104'
    configPath = curPath+'\\config\\'
    outputPath = curPath + '\\output\\seg_data\\'
    endlineLable = "\r\n" # row tag for windows
    globalFile = configPath + 'global_win.cfg'
else:
    configPath = curPath+'/config/'
    outputPath = curPath + '/output/seg_data/'
    endlineLable = "\n"   # row tag for linux
    globalFile = configPath + 'global.cfg'
config = ConfigParser.ConfigParser()
config.readfp(open(globalFile,'rb'))
FILEPATH = config.get("global","pmFilePath")

print ('FILEPATH is %s' % FILEPATH)
if FILEPATH == '':
    print('need to configure PM file path in '+ configPath + 'global.cfg')
    sys.exit()

CFGPATH = configPath + 'chart_seg.cfg'
fileNameList = []
daylastfile = []
g_ModuleGroupList = []    # store module-statName|group

valueStr = re.compile('.+measValue.+')

#configuration by user-defined
partLable = ("<",">")
sectionLable = ("{","}")
displaynameGroupSep = '|'
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
    
#Get all counters for a group under a module
def parseCFGByUserDefined (cfgFile, tag):
    f = open(cfgFile,"rb")
    strFileContent = f.read()
    f.close()
    vardict = {}
    var1 = getSectionMap(strFileContent)
    for k,v in var1.items():
        vardict[k] = v
    return vardict[tag]

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
"""
Get Module name and group name in configuration file
Output: ModuleName-sheetName|GroupName
"""
def getPlatformAndSection (cfgFile):
    global g_ModuleGroupList;
    
    f = open(cfgFile,"rb")
    strFileContent = f.read()
    f.close()
    
    var1 = getPlatformMap(strFileContent)
    g_ModuleGroupList = []
    for k in var1.items():
        g_ModuleGroupList.append(k[0])


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
    radiusInstance = []
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
                if temp[0] == 'AaaInterface' and len(temp) == 4 and temp[1] == 'statisticsPerServer':
                    aaaInstance.append(temp[3])
                if temp[0] == 'AaaInterface' and len(temp) == 4 and temp[1] == 'statisticsPerRadiusServer':
                    radiusInstance.append(temp[3])
                if temp[0] == 'Gtp' and len(temp) == 4 and temp[1] == 'statisticsPerPgw':
                    gtpInstance.append(temp[3])
                if temp[0] == 'Gtp' and len(temp) == 4 and temp[1] == 'QCI':
                    qciInstance.append(temp[3])
        else:
            pass
    return aaaInstance, gtpInstance, qciInstance, radiusInstance

##################
  
def writeExcel(xlsFileName):
    
    global fileNameList

    #create excel work sheet
    book=xlwt.Workbook()
    
    groupItem = "G1"

    for sheetname in g_ModuleGroupList:            
        groupName = sheetname
        fields = parseCFGByUserDefined(CFGPATH, sheetname)
        fields = fields.split('\n')
        
        if 'SystemManagement' in sheetname:
            sheetname = sheetname.replace('SystemManagement','SysMgmt')
        
        if len(sheetname) > 31:
            sheetname = sheetname[0:30]    	
        sheet = book.add_sheet(groupItem +'_'+ sheetname, cell_overwrite_ok=True)
        sheet.write(0,0,"Group ID")
        sheet.write(0,1,"Time")
        
        i = 1
        z = 1

        #print ('collecting the %s %s data' % (groupItem,sheetname))
        #print 'fileNameList = %s' % fileNameList
        for fileName in fileNameList:
            #print 'Collecting the %s %s data from the %sth file --%s' % (groupItem, sheetname, i, (FILEPATH + fileName))
            print 'Collecting the %s %s data from file --%s' % (groupItem, sheetname,(FILEPATH + fileName))
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

            if groupID != groupItem:
                continue
            else:
                aaaDict = {}    # Server=xxx : values(<r p="1"></r>)
                aaaObjDict = {} # Module-Group : aaaDic
                radiusDict = {}
                radiusObjDict = {}
                sysDict = {}
                copyAAA = []
                copyRadius = []
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

                    for measObj in tempValueList:
                        measObjLdn = measObj.attrib
                        instance = measObjLdn.get('measObjLdn')
                        temp = instance.split(',')
                        tempObjLdn = temp[2][0:-2]
                        
                        if len(temp) == 4:
                            if temp[2] == 'Ipsec=1' and temp[3][0:4] == 'Vpn=':
                                copyAAA = measTypeList[:]
                                value = measObj.getchildren()
                                aaaDict[temp[3]] = value
                                aaaObjDict = {tempObjLdn:aaaDict}
                            elif temp[0] == 'AaaInterface' and temp[1] == 'statisticsPerRadiusServer':
                                copyRadius = measTypeList[:]
                                value = measObj.getchildren()
                                radiusDict[temp[3]] = value
                                radiusObjDict = {tempObjLdn:radiusDict}
                            elif temp[2] == 'SystemManagement=1' and temp[3][2:6] == 'CpCpu=':
                                copySYS = measTypeList[:]
                                value = measObj.getchildren()
                                sysDict[temp[3]] = value
                                sysObjDict = {tempObjLdn:sysDict}
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
                #a is used to indicate the count of Delta 
                a = 0
                #j = 2
                for dictItem in tempList:
                    #print ('====dictItem.keys()[0]==sheetname===%s,%s' % (dictItem.keys()[0],sheetname))
                    record = []
                    #j=2
                    if dictItem.keys()[0] == groupName:
                        for field in fields:
                            ##Delta means (current value - last value)
                            if field == '':
                                continue
                            #print ('j is %d and field is %s' % (j,field))
                            if 'Delta' in field:
                                field = field.split('--')[0]
                                if i == 1:
                                    sheet.write(0,j,field)
                                    sheet.write(i,j,(dictItem.values()[0]).get(field))
                                    sheet.write(i,j,0)
                                    sheet.write(i,0,groupID)
                                    sheet.write(i,1,date)
                                    record.append((dictItem.values()[0]).get(field))
                                else:
                                    #get the last cc data of last file
                                    result = config.get('record','lastrecord')
                                    #print ('aaaa--result = %s' % result)
                                    temp = int((dictItem.values()[0]).get(field)) - int(result[a])
                                    if temp < 0:
                                        temp = 0
                                    sheet.write(i,j,temp)
                                    sheet.write(i,0,groupID)
                                    sheet.write(i,1,date)
                                    record.append((dictItem.values()[0]).get(field))
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
                    config.write(open(globalFile,'w'))      
                    #print 'record = %s' % record

                #data process for AAA
                if len(aaaObjDict) != 0:
                    if groupName == (aaaObjDict.keys()[0]): 
                        #print ('aaaObjDict.keys()[0])[0:30]==========%s' % (aaaObjDict.keys()[0])[0:30])
                        if i == 1:                                   
                            sheet.write(0,2,"Instance")
                        aaaValueDict = {}
                        aaaItem = aaaDict.items()
                        #print '========aaaItem= %s' % aaaItem
                        record = []
                        #b is used to record the number of Delta 
                        b = 0
                        for elem in aaaItem:           
                            temp = list(elem)
                            for aaavalue in temp[1]:
                                for aaaField in copyAAA:
                                    if aaaField.attrib.values()[0] == aaavalue.attrib.values()[0]:
                                        aaaValueDict[aaaField.text] = aaavalue.text
                            #print '=======aaaValueDict = %s' %  aaaValueDict           
                            x = 3
                            #the second column of the worksheet is "Instance"
                            #print '=======groupID = %s' %  groupID
                            sheet.write(z,0,groupID)
                            sheet.write(z,1,date)
                            sheet.write(z,2,temp[0])
                            
                            #get AAA data
                            for field in fields:
                                #print '=======field = %s' % field
                                if 'Delta' in field:
                                    #print ('cc field === %s' % field)
                                    field = field.split('--')[0]
                                    record.append(aaaValueDict.get(field))
                                    #the first pm file's datas are assigned value 0
                                    if i == 1:
                                        if aaaValueDict.has_key(field):
                                            #only one time for add field name
                                            if z == 1 and i == 1:
                                                sheet.write(0,x,field)
                                                sheet.write(z,x,0)
                                            else:
                                                sheet.write(z,x,0)
                                                
                                        x += 1
                                    else:
                                        #get the last data from config file
                                        result = config.get('record','lastrecord')
                                        if aaaValueDict.has_key(field):
                                            #next data subtract last data
                                            #print ('aaaValueDict_field %s, result_b %s' % (aaaValueDict.get(field),result[b]))
                                            temp = int(aaaValueDict.get(field)) - int(result[b])
                                            if temp < 0:
                                                temp = 0
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
                        #print 'record recordrecordrecordrecord= %s' % record
                        #record the last data in config file
                        if len(record) != 0:
                            config.set('record','lastrecord',record)
                            config.write(open(globalFile,'w')) 
                
                #data process for Radius
                if len(radiusObjDict) != 0:
                    if groupName == (radiusObjDict.keys()[0]): 
                        if i == 1:                                   
                            sheet.write(0,2,"Instance")
                        radiusValueDict = {}
                        radiusItem = radiusDict.items()
                        #print '========radiusItem= %s' % radiusItem
                        record = []
                        #b is used to record the number of Delta 
                        b = 0
                        for elem in radiusItem:           
                            temp = list(elem)
                            for radiusvalue in temp[1]:
                                for radiusField in copyRadius:
                                    if radiusField.attrib.values()[0] == radiusvalue.attrib.values()[0]:
                                        radiusValueDict[radiusField.text] = radiusvalue.text
                            #print '=======radiusValueDict = %s' %  radiusValueDict           
                            x = 3
                            #the second column of the worksheet is "Instance"
                            #print '=======groupID = %s' %  groupID
                            sheet.write(z,0,groupID)
                            sheet.write(z,1,date)
                            sheet.write(z,2,temp[0])
                            
                            #get Radius data
                            for field in fields:
                                #print '=======field = %s' % field
                                if 'Delta' in field:
                                    #print ('cc field === %s' % field)
                                    field = field.split('--')[0]
                                    record.append(radiusValueDict.get(field))
                                    #the first pm file's datas are assigned value 0
                                    if i ==1:
                                        if radiusValueDict.has_key(field):
                                            #only one time for add field name
                                            if z == 1 and i == 1:
                                                sheet.write(0,x,field)
                                                sheet.write(z,x,0)
                                            else:
                                                sheet.write(z,x,0)
                                                
                                        x += 1
                                    else:
                                        #get the last data from config file
                                        result = config.get('record','lastrecord')
                                        if radiusValueDict.has_key(field):
                                            #next data subtract last data
                                            temp = int(radiusValueDict.get(field)) - int(result[b])
                                            if temp < 0:
                                                temp = 0
                                            if z == 1 and i == 1:
                                                sheet.write(0,x,field)
                                                sheet.write(z,x,temp)
                                            else:
                                                sheet.write(z,x,temp)
                                        x += 1
                                    b += 1    
                                else:
                                    if radiusValueDict.has_key(field):
                                        #only one time for add field name
                                        if z == 1 and i == 1:
                                            sheet.write(0,x,field)
                                            sheet.write(z,x,radiusValueDict.get(field))
                                        else:
                                            sheet.write(z,x,radiusValueDict.get(field))
                                    x += 1
                            z += 1
                        #print 'record recordrecordrecordrecord= %s' % record
                        #record the last data in config file
                        if len(record) != 0:
                            config.set('record','lastrecord',record)
                            config.write(open(globalFile,'w'))
                    
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

def printLog(obj):
        print "%%%%%%%%%%%%%%%%%%%%%%%%"
        print obj
    
if __name__ == '__main__':

    FILEPATH = '../../../pdctool/pilotpm/'
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
    mkdirs(outputPath)
    excelName = date + '.' + hhmm + str(timezone) + '_' + hostname + '_source_data.xls'
    excelName = os.path.join(outputPath,excelName)
    
    # initialize module-displayname|group
    getPlatformAndSection(CFGPATH)
    
    writeExcel(excelName)
        
    print ('data analyzing......')
    #get aaa and gtp instance count
    aaaInstance, gtpInstance, qciInstance, radiusInstance = getGtpAndAaaInstance(fileNameList[0])
    aaaInstance = makeListElementUnique(aaaInstance)
    radiusInstance = makeListElementUnique(radiusInstance)
    gtpInstance = makeListElementUnique(gtpInstance) 
    qciInstance = makeListElementUnique(qciInstance)
    #print ('aaaInstance === %s' % aaaInstance)
    #print ('gtpInstance === %s' % gtpInstance) 
    #print ('qciInstance === %s' % qciInstance)
    
    workbook = xlsxwriter.Workbook(os.path.join(outputPath,date+'-'+hhmm+timezone+'_'+hostname+'_wmg_graph.xlsx'))
    
    #data = xlrd.open_workbook('/home/mjw/pdctool/output/cp04_or_wmg_data/20150519.1505+0800_ubuntu_source_data.xls')
    data = xlrd.open_workbook(excelName)
    sheets = data.sheet_names()
    date_format = workbook.add_format({'num_format': 'dd/mm/yyyy hh:mm'})   
    
    aaaSheets = []
    gtpSheets = []
    qciSheets = []
    radiusSheets = []
    radiusstr = re.compile('.+RadiusSvr.*')
    aaastr = re.compile('.+AaaInterface-Dia.+')
    gtpstr = re.compile('.+Gtp-Pgw.+')
    qcistr = re.compile('.+Gtp-QCI')
    sheetNamePart = ['SysMgmt-cpuAndMem','SysMgmt-throughput_M','SysMgmt-throughput_K','Sc-Total', 'Sc-Active','Sc-Emergency_Call',\
        'Sc-Handoff_Call', 'Sc-Call_Duration','IpStack-Pkg','IpStack-Bytes','IpStack-FragmentReassemble',\
        'AaaInterface-Phase_Start', 'AaaInterface-Phase_Final', 'AaaInterface-Radius_Client', 'AaaInterface-Radius_Proxy',\
        'Ipsec-General', 'Ipsec-Message', 'Ipsec-Kpi', 'Ipsec-Ceps','Dns-General','Gtp-Bears','Gtp-Tunnels',\
        'Twag-TotalSecure', 'Twag-ActiveSecure', 'Twag-TotalOpen', 'Twag-ActiveOpen', 'Twag-Ceps']
    print ('sheet sheets is %s' % sheets)
    
    for namePart in sheetNamePart:
        allworksheet = workbook.add_worksheet(namePart)
        chartPosition = 1
        flag = 0
        for sheet in sheets:
            if aaastr.match(str(sheet)):
                aaaSheets.append(sheet)
            elif radiusstr.match(str(sheet)):
                radiusSheets.append(sheet)
            elif gtpstr.match(str(sheet)):
                gtpSheets.append(sheet)
            elif qcistr.match(str(sheet)):
                qciSheets.append(sheet)
            elif namePart in sheet:
                print ('sheet name is %s' % sheet)
                chart = workbook.add_chart({'type': 'line'})
                dataSheetName = "g"+sheet[1:]
                worksheet = workbook.add_worksheet(dataSheetName)
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
                    chart_series(dataSheetName, nrows-1, cur_row)
        
                chart.set_size({'width': 700, 'height': 350})
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
                else: 
                    allworksheet.insert_chart(chartPosition, 11, chart)
                    chartPosition += 20
                flag += 1
                
                
    
    ##############################################
    #deal with Sc_Call_Distribution, it only need the last data of every day
    allDistribution = workbook.add_worksheet('Sc-Call_Distribution')
    DistributionChartPosition = 1
    DistributionFlag = 0
    for item in sheets:        
        if str(item)[3:] == 'Sc-Call_Distribution':
            #print ('Sc-Call_Distribution ===== %s' % item)
            disSheetName = 'g'+'_'+str(item)
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
                        tempList.append(cell_value)
                    
                      
                #tempList.sort()
                #print ('tempList ===== %s' % tempList)
                if len(tempList) != 0 :
                    lastDateList.append(tempList[-1])
            #print ('lastDateList ===== %s' % lastDateList)
            
            c = 1
            for lastDate in lastDateList:
                for i in range(1,nrows):
                    cell_value = table.cell_value(i,1)                    
                    if lastDate == cell_value:
                        for j in range(ncols):
                            cell_value = table.cell_value(i,j,)
                            if type(cell_value) == unicode and cell_value.isdigit():
                                worksheet.write(c,j,string.atoi(cell_value))
                            else:
                                worksheet.write(c,j,cell_value)
                        #worksheet.write_row(c,0,table.row_values(i))
                        c += 1
            
            for cur_row in range(2,ncols):
                    chart_series(disSheetName, c-1, cur_row)
        
            chart.set_size({'width': 700, 'height': 350})
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
            else: 
                allDistribution.insert_chart(DistributionChartPosition, 11, chart)
                DistributionChartPosition += 20
            DistributionFlag += 1            
                
    #deal with Sc_Call_Distribution end
    ##############################################
    
    aaaSheets =  makeListElementUnique(aaaSheets) 
    gtpSheets =  makeListElementUnique(gtpSheets) 
    qciSheets =  makeListElementUnique(qciSheets)
    radiusSheets = makeListElementUnique(radiusSheets)
    #print ('aaaSheet ==== %s' % aaaSheets)
    #print ('gtpSheet ==== %s' % gtpSheets) 
    #print ('qciSheets ==== %s' % qciSheets)
    #sum_sheet_wb = workbook.add_worksheet('sum_sheet')
    tmpSheetList = []
    
    ################################################################################       
    #process for QCI
    graphSheet_Qci = workbook.add_worksheet("Gtp-QCI")
    chartPosition = 1
    flag = 0
    for qciSheet in qciSheets:
        chart = workbook.add_chart({'type': 'line'})
        dataSheetName = "g" + qciSheet[1:]
        dataSheet = workbook.add_worksheet(dataSheetName)
        table = data.sheet_by_name(qciSheet)
        nrows = table.nrows
        ncols = table.ncols 
       
        # GroupID and time
        for col in range(0, 2):
            cell_value = table.cell_value(0, col)
            dataSheet.write(0, col, cell_value)
        
        #Add counters for every QCI
        col_num = 2
        for instance in qciInstance:
            instanceStr = instance.split("=")
            instanceName = instanceStr[1]
            for col in range(3, ncols):
                counterName = table.cell_value(0, col)
                counterName = instanceName + " " + counterName
                dataSheet.write(0, col_num, counterName)
                col_num += 1
        
        rowIndex = 0
        colIndex = 2
        lastTime = ""
        for i in range(1, nrows):
            currentTime = table.cell_value(i, 1)
            if currentTime != lastTime:
                rowIndex +=1
                colIndex = 2
                dataSheet.write(rowIndex, 0, table.cell_value(i, 0))
                dataSheet.write(rowIndex, 1, table.cell_value(i, 1))
                lastTime = currentTime
            # write value of counters
            for j in range(3, ncols):
                cell_value = table.cell_value(i, j)
                if type(cell_value) == unicode and cell_value.isdigit():
                    dataSheet.write(rowIndex, colIndex, string.atoi(cell_value))
                else:
                    dataSheet.write(rowIndex, colIndex, cell_value)
                colIndex += 1           
                      
        for cur_col in range(2, len(qciInstance) * (ncols-3) + 2):
            chart_series(dataSheetName, rowIndex, cur_col)        
    
        chart.set_size({'width': 700, 'height': 350})
        chart.set_title ({'name': qciSheet + '_stat.'})
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
            
        if flag % 2 == 0:
            graphSheet_Qci.insert_chart(chartPosition, 0, chart) 
        else: 
            graphSheet_Qci.insert_chart(chartPosition, 11, chart)
            chartPosition += 20
        flag += 1
    
    for aaaSheet in aaaSheets:
        preName = aaaSheet[0:2]
        sheetnamestr = aaaSheet.split('-')
        sheet_for_aaa = workbook.add_worksheet(aaaSheet[0:2]+'_'+sheetnamestr[1])
        table = data.sheet_by_name(aaaSheet)
        nrows = table.nrows
        ncols = table.ncols
        #print ('sheet ncols is %s' % ncols) 
        #print ('sheet nrows is %s' % nrows)
        chartPosition = 1
        flag = 0
        for instance in aaaInstance:
            partInstance = instance.split('=')
            copySheetName = 'g'+aaaSheet[1:2]+'_'+sheetnamestr[1]+'_'+partInstance[1]
            copySheetName = copySheetName[0:31]
            worksheet = workbook.add_worksheet(copySheetName)
            worksheet.write_row(0,0,table.row_values(0))
            chart = workbook.add_chart({'type': 'line'})
            i = 1
            for rownum in range(1, nrows):
                if instance == table.cell(rownum,2).value:
                    #print ('i === %s' % i)
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
                chart_series_aaagtp(copySheetName, nrows/len(aaaInstance)-1, cur_row)
                
            chart.set_size({'width': 700, 'height': 350})
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
            #sheet_for_aaa.insert_chart(chartPosition, 0, chart)
            #chartPosition += 16
            if flag % 2 == 0:
                sheet_for_aaa.insert_chart(chartPosition, 0, chart) 
            else: 
                sheet_for_aaa.insert_chart(chartPosition, 11, chart)
                chartPosition += 20
            flag += 1
            
    # process for Radius
    for radiusSheet in radiusSheets:
        preName = radiusSheet[0:2]
        sheetnamestr = radiusSheet.split('-')
        sheet_for_radius = workbook.add_worksheet(radiusSheet[0:2]+'_'+sheetnamestr[1])
        table = data.sheet_by_name(radiusSheet)
        nrows = table.nrows
        ncols = table.ncols
        #print ('sheet ncols is %s' % ncols) 
        #print ('sheet nrows is %s' % nrows)
        chartPosition = 1
        flag = 0
        for instance in radiusInstance:
            partInstance = instance.split('=')
            copySheetName = 'g'+radiusSheet[1:2]+'_'+sheetnamestr[1]+'_'+partInstance[1]
            copySheetName = copySheetName[0:31]
            worksheet = workbook.add_worksheet(copySheetName)
            worksheet.write_row(0,0,table.row_values(0))
            chart = workbook.add_chart({'type': 'line'})
            i = 1
            for rownum in range(1, nrows):
                if instance == table.cell(rownum,2).value:
                    #print ('i === %s' % i)
                    for colnum in xrange(ncols):
                        cell_value = table.cell_value(rownum,colnum,)
                        if type(cell_value) == unicode and cell_value.isdigit():
                            worksheet.write(i, colnum, string.atoi(cell_value))
                        else:
                            worksheet.write(i,colnum,cell_value)

                    i += 1
 
            #tmpSheetList.append(radiusSheet+instance)
            tmpSheetList.append(worksheet)
            #print ('tmpSheetList is %s' % tmpSheetList)            
            for cur_row in range(3,ncols):
                chart_series_aaagtp(copySheetName, nrows/len(radiusInstance)-1, cur_row)
                
            chart.set_size({'width': 700, 'height': 350})
            chart.set_title ({'name': radiusSheet[3:]+'_'+instance})
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
            #sheet_for_radius.insert_chart(chartPosition, 0, chart)
            #chartPosition += 16
            if flag % 2 == 0:
                sheet_for_radius.insert_chart(chartPosition, 0, chart) 
            else: 
                sheet_for_radius.insert_chart(chartPosition, 11, chart)
                chartPosition += 20
            flag += 1

    ################################################################################       
    #process for GTP
    for gtpSheet in gtpSheets:
        preName = gtpSheet[0:2]
        sheetnamestr = gtpSheet.split('-')
        sheet_for_gtp = workbook.add_worksheet(preName+'_'+sheetnamestr[1])
        table = data.sheet_by_name(gtpSheet)
        nrows = table.nrows
        ncols = table.ncols
        chartPosition = 1
        flag = 0
        for instance in gtpInstance:
            partInstance = instance.split('=')
            if partInstance[1] == 'ffff:ffff:ffff:ffff:ffff:ffff:ffff:ffff':
                sheetname = 'g'+gtpSheet[1:2]+'_'+sheetnamestr[1]+'_'+'ffff.ffff.ffff.ffff.ffff.ffff.ffff.ffff'
                sheetname = sheetname[0:31]
            else:
                sheetname = preName+'_'+sheetnamestr[1][0:6]+'_'+instance
            if len(sheetname) > 31 :
                sheetname = sheetname[0:31]
            sheetname = "g" + sheetname[1:]

            worksheet = workbook.add_worksheet(sheetname)
            worksheet.write_row(0,0,table.row_values(0))
            chart = workbook.add_chart({'type': 'line'})
            i = 1
            for rownum in range(1, nrows):
                if instance == table.cell(rownum,2).value:
                    #print ('gtp=== i === %s' % i)
                    for colnum in xrange(ncols):
                        cell_value = table.cell_value(rownum,colnum,)
                        if type(cell_value) == unicode and cell_value.isdigit():
                            worksheet.write(i, colnum, string.atoi(cell_value))
                        else:
                            worksheet.write(i,colnum,cell_value)

                    i += 1 
            #print ('nrows is %s' % nrows)            
            for cur_row in range(3,ncols):
                chart_series_aaagtp(sheetname, nrows/len(gtpInstance)-1, cur_row)
                
            chart.set_size({'width': 700, 'height': 350})
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
            #sheet_for_gtp.insert_chart(chartPosition, 0, chart)
            #chartPosition += 16
            if flag % 2 == 0:
                sheet_for_gtp.insert_chart(chartPosition, 0, chart) 
            else: 
                sheet_for_gtp.insert_chart(chartPosition, 11, chart)
                chartPosition += 20
            flag += 1
    
    print ('data charts are generating......')
    
    #hide the worksheets
    worksheets = workbook.worksheets()
    for worksheet in worksheets:
        if worksheet.get_name().startswith("g"):
            worksheet.hide()
    
    workbook.close()
    
    print ('Collect data successfully.')

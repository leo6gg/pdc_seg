#!/usr/bin/python
# -*- coding: utf-8 -*-
from xml.etree import ElementTree as ET
import re
import csv
import os
import time
import sets
import os.path
import datetime
import ConfigParser
import sys

fileTime = ''
timezone = ''

measInfoList = []
fileNameList = []

curPath = os.getcwd()
tmpPath = os.path.abspath(os.path.join(curPath,os.path.pardir))
workPath = os.path.abspath(os.path.join(tmpPath,os.path.pardir))
configPath = curPath+'/config/'
#print configPath
globalFile = configPath + 'global.cfg'
#print ('globalFile == %s' % globalFile)
config = ConfigParser.ConfigParser()
config.readfp(open(globalFile,'rb'))
FILEPATH = config.get("global","pmFilePath")
if FILEPATH == '':
    print('need to configure PM file path in '+ configPath + 'global.cfg')
    sys.exit()

#FILEPATH = '/home/mjw/Files/cp04/'
#FILEPATH = '/home/mjw/Files/15b/'
PDCDATAPATH = curPath+'/output/wmg_data/'
DATABAKUPPATH = PDCDATAPATH + 'PMFilesDataBakup/'
ONDEMANDPATH = PDCDATAPATH + 'on_demand/'
TEMPPATH = PDCDATAPATH + 'temp/tmp/'
ARCHIVEPATH = '/md/services/wmg/pdc/archive/'

valueStr = re.compile('.+measValue.+')


###########################################
##sourceDir:
##targetDir:
###########################################
def copyFiles (sourceDir, targetDir):
    for file in os.listdir(sourceDir):
        sourceFile = os.path.join(sourceDir,  file)
        targetFile = os.path.join(targetDir,  file)
        if os.path.isfile(sourceFile): 
            if not os.path.exists(targetDir):
                os.makedirs(targetDir)
            if not os.path.exists(targetFile) or(os.path.exists(targetFile) and (os.path.getsize(targetFile) != os.path.getsize(sourceFile))):
                open(targetFile, "wb").write(open(sourceFile, "rb").read())
        if os.path.isdir(sourceFile):
            First_Directory = False
            copyFiles(sourceFile, targetFile)


#make the element in list unique
def unique(L):
    #return [x for x in L if x not in locals()['_[1]']]
    result = list(sets.Set(L))
    return result


def parseXML(fileName):
    measTypeList = []
    measValueList = []
    doc = ET.ElementTree(file=fileName)
    root = doc.getroot()
    all_node = root.getchildren()
    measData = all_node[1]
    children = measData.getchildren()
    measInfos = children[1:]
    for item in measInfos:
        measInfoTmp = item.getchildren()[2:]
        for item1 in measInfoTmp:
            if valueStr.match(str(item1)):
                measValueList.append(item1)
            else:
                measTypeList.append(item1.text)
    return measValueList,measTypeList,measInfos

def getAAAandGtpCount(fileName):
    doc = ET.ElementTree(file=fileName)
    root = doc.getroot()
    all_node = root.getchildren()
    measData = all_node[1]
    children = measData.getchildren()
    measInfos = children[1:]
    aaaCount = 0
    gtpCount = 0
    aaaInstanceList = []
    gtpinstanceList = []
    for item in measInfos:
        measTypeList = []
        measValueList = []
        measInfoTmp = item.getchildren()[2:]
        for item1 in measInfoTmp:
            if valueStr.match(str(item1)):
                measValueList.append(item1)
            else:
                measTypeList.append(item1.text)
        for it in measValueList:
            measObjLdnValue = it.attrib['measObjLdn']
            temp = measObjLdnValue.split(',')
            if temp[1] == 'Aaa':
                aaaCount += 1
                aaaInstanceList.append(it)
            elif temp[1] == 'Gtp' and len(temp) == 4:
                gtpCount += 1
                gtpinstanceList.append(it)
    return aaaInstanceList, gtpinstanceList

def getValues(valueItem, fileName, resultList):
    valueList = []
    value = valueItem.getchildren()
    for valueBin in value:
        valueList.append(valueBin.text)
        valueList.insert(0,fileName)
        fileTime = fileName.split('_')
        valueList.insert(1,fileTime[0][1:])
        valueList.insert(2,fileTime[1]+'_')
        valueTuple = tuple(valueList)
        resultList.append(valueTuple)
    return resultList
   

###########################################
##
##
###########################################
def createDataFile(dict, group, CSVPath, fileSuffix,fieldNameList):
    tempList = []
    csvFile = ''
    for item in dict:
        if item.keys()[0] == str(group):
            tempList.append(item.values()[0])
            csvFile = file(CSVPath+'G'+str(group)+fileSuffix,'wb')
    #print 'create file name is %s' % csvFile
    writer = csv.writer(csvFile, delimiter='|', dialect='excel')
    tempTuple = tuple(fieldNameList)
    tempList.insert(0,tempTuple)
    writer.writerows(tempList)
    csvFile.close()
 
###########################################
##FILEPATH: path of PM file 
##CSVPath: path for generated CSV file
###########################################
def process(FILEPATH, CSVPath):

    measTypeDictList = []
    #get each module fields.
    measValueList, measTypeList, measInfos = parseXML(FILEPATH+fileNameList[0])
    for item in measInfos:
        measTypeDict = {}
        measInfoTmp = item.getchildren()[2:]
        measTypeList = []
        measValueList = []
        for item1 in measInfoTmp:
            if valueStr.match(str(item1)):
                measValueList.append(item1)
            else:
                if item1.text == 'Reset Time':
                    pass
                else:
                    measTypeList.append(item1.text)

        #for all 
        measTypeList.insert(0,'Time')
        for valueItem in measValueList:
            valueList = []
            measObjLdnValue = valueItem.attrib['measObjLdn']
            #print('======%s' % measObjLdnValue)
            temp = measObjLdnValue.split(',')
            if len(temp) == 4:
                moduleName = 'G'+temp[2][-1]+'_'+temp[0]+'_'+temp[1]+'_'+temp[3]
            else:
                moduleName = 'G'+temp[2][-1]+'_'+temp[0]+'_'+temp[1]
                
            #measTypeList.insert(0,'Time')
            #measTypeDict compose of measObjLdn infor and measType    
            measTypeDict = {moduleName:measTypeList}
            measTypeDictList.append(measTypeDict)

        #print('+++++++measTypeList = %s' % measTypeList)
    #print('++++length++of++measTypeDictList = %s' % len(measTypeDictList))
    #print('+++++++measTypeDictList = %s' % measTypeDictList)
        
    global fileTime
    global timezone

    groupIDList = []

    i = 1
    dict = {}
    valueDictList = []
    #re match reset time
    reStr = re.compile('\d{1,2}/\d{1,2}/\d{4}\s+\d{1,2}:\d{1,2}:\d{1,2}\s.+')
    for fileName in fileNameList:

        print 'Collect the %sth file --%s' % (i, (FILEPATH + fileName))
        valueList1, typeList1, measInfos = parseXML(FILEPATH+fileName)
        fileTime = fileName.split('_')
        timezone = fileTime[0][-5:]

        groupID = fileTime[1].split('-')[-1]
        groupIDList.append(groupID)

        #Date
        nameItem = fileName.split('_')
        tt = nameItem[0].split('.')
        dateStr = tt[0][1:]
        yy = dateStr[:4]
        mm = dateStr[4:6]
        dd = dateStr[6:8]
        hh = tt[1][10:12]
        MM = tt[1][12:14]
        ss = '00'
        date = yy+'-'+mm+'-'+dd+' '+hh+':'+MM+':'+ss
               
        #print('length for valueList1 = %s' % len(valueList1))
        for valueItem in valueList1:
            valueList = []
            measObjLdnValue = valueItem.attrib['measObjLdn']
            #print('======%s' % measObjLdnValue)
            temp = measObjLdnValue.split(',')
            if len(temp) == 4:
                moduleName = 'G'+temp[2][-1]+'_'+temp[0]+'_'+temp[1]+'_'+temp[3]
            else:
                moduleName = 'G'+temp[2][-1]+'_'+temp[0]+'_'+temp[1]
            value = valueItem.getchildren()
            #print('===value===%s' % value)
            for valueBin in value:
                #ignore Reset time
                if reStr.match(valueBin.text):
                    pass
                else:
                    valueList.append(valueBin.text)
            valueList.insert(0,date)
            #print('===valueList===%s' % valueList)
            #dict compose of measObjLdn infor and valueList
            dict = {moduleName:valueList}
            valueDictList.append(dict)
    
        i += 1
    #print('===valueDictList===%s' % valueDictList)
    #print '+++++sysMgmtListALL=%s' % sysMgmtListALL
    #print '+++++AaaList=%s' % AaaList
    group = unique(groupIDList)
    print ('+++++group=%s' % group)
    
    #create CSV file
    csvFile = ''
    for groupItem in group:
        for item in measTypeDictList:
            tempList = []
            csvFileName = 'G'+groupItem+(item.keys()[0])[2:]
            csvFile = file(CSVPath+csvFileName+'.log','wb')
            writer = csv.writer(csvFile, delimiter='|', dialect='excel')
            #print ('@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@')
            #print ('fields keys %s' % item.keys()[0])
            for item1 in valueDictList:
                #print ('fields keys in values %s' % item1.keys()[0])
                if csvFileName == (item1.keys()[0]):
                    tempList.append(item1.values()[0])
            tempList.insert(0,item.values()[0])
            writer.writerows(tempList)    
    #print 'create file name is %s' % csvFile    
    csvFile.close()
            

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


###########################################
##source: The files or directories you want to compress
##target: The target directory you want to place the tar file
###########################################
def backupFiles (source, target):
    global timezone
    #today=target+time.strftime('%Y-%m-%d')
    today=target
    now=time.strftime('%H%M')
    mmm = time.strftime('%b')
    hostnameTmp = os.popen('hostname')
    hostname = hostnameTmp.readline()
    hostname_len = len(hostname)
    if hostname_len > 11:
        hostname = hostname[0:11]
        if '.' in hostname:
            hostname = hostname.replace('.','-')
    #os.sep = /
    #now = 071350
    # if the path not exist, then create it
    if not os.path.exists(today):
        os.makedirs(today)
        print('Successfully created directory',today)

    #every first day of month at 02:00 AM
    firstDay = time.strftime('%d%H%M')
    if firstDay == '010200':
        #monthly package
        tempTarget=today +os.sep+mmm+'_'+hostname.strip()+'_'+time.strftime('%Y-%m')+timezone+'monthly_wmg_pdc'+'.tar.gz'
    else:
        #on-demand package
        tempTarget=today +os.sep+mmm+'_'+hostname.strip()+'_'+time.strftime('%Y%m%d')+'-'+now+timezone+'_wmg_pdc'+'.tar.gz'
    tar_command = 'cd %s;cd ../;tar -zcf %s ./' %(source,tempTarget)
    #tar_command = 'cd %s;tar -zcf %s ./' %(source,tempTarget)
    if os.system(tar_command) == 0:
        print ('Successful backup to', tempTarget)
    else:
        print('Backup Failed')


if __name__ == '__main__':

    print 'Begin to collect data for epdg, please waiting...'
    dateList = []
    for root, dirs, files in os.walk(FILEPATH):
        for fn in files:
            #print 'files are: %s' % fn
            fileNameList.append(fn)
    fileNameList.sort()
    for item in fileNameList:
        dateStrList = item.split('.')
        dateStr = dateStrList[0][1:]
        dateList.append(dateStr)
    date = unique(dateList)
    #print 'date = %s' % date
    
    path = mkdirs(PDCDATAPATH)
    bakuppath = mkdirs(DATABAKUPPATH+time.strftime('%Y-%m-%d'))
    #print 'bakuppath =======%s' % bakuppath
    process(FILEPATH,bakuppath + os.sep)
    #backupFiles(bakuppath + os.sep, TEMPPATH)

    for item in date:
        targetDir = mkdirs(TEMPPATH +item)
        copyFiles (bakuppath+os.sep, targetDir)
        for file in os.listdir(targetDir):
            f = open(targetDir + os.sep + file, "r")
            content = f.readlines()
            #print 'content = %s' % content
            for row in content[1:]:
                timeStr = row.split('|')
                dateStr = timeStr[0].split(' ')
                contentDate = dateStr[0].replace('-','')
                hour = dateStr[1].split(':')[0]
                min = dateStr[1].split(':')[1]
                #print 'date111111111 = %s' % date ONDEMANDPATH
                if contentDate != item :
                    content.remove(row)
                if contentDate == item and hour == '00' and min == '00':
                    day = datetime.datetime.strptime(timeStr[0],'%Y-%m-%d %H:%M:%S') + datetime.timedelta(days=1)
                    newDate = day.strftime('%Y-%m-%d %H:%M:%S')
                    #print 'row ============== %s' % row
                    content[content.index(row)] = newDate + row[19:]
            #for Sc call_Distribution only need the last data.
            Call_Distribution = file.split('_')
            if Call_Distribution[2] == 'Call' and Call_Distribution[3] == 'Distribution.log':
                rowName = content[0]
                lastRow = content[-1]
                content = rowName + lastRow
                #print ('content[-1] = %s' % content)
            
            new_f = open(targetDir + os.sep + file, "w")
            new_f.writelines(content)
            f.close()
            new_f.close()

    backupFiles(TEMPPATH, ONDEMANDPATH)

    print 'Collect data successfully.'

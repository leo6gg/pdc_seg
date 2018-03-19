#!/usr/bin/python3
import sys
sys.path.append("/usr/lib/pexpect/")
import pexpect
import time
import os
import re

epdgpm = "/md/epdg/pm/"
wmgpm = "/md/wmg/pm/"
wmgPdcArchive = "/md/wmg/pdc/archive/temp/tmp/"
epdgPdcArchive = "/md/epdg/pdc/archive/temp/tmp/"
open_file = ""

releaseList = ['1r26b03','1r26b04','1r26b05','1r26b06']

hostname = (os.popen("hostname")).readline().strip()

#access ssr shell
child = pexpect.spawn("ssh -q -o UserKnownHostsFile=/dev/null -o StrictHostKeyChecking=no -o ForwardX11=no dev@%s" % hostname, timeout = 300*300)
result = child.expect([hostname + "#", "password:"])

print ("result = %s" % result)
if result == 1:
    child.sendline("dev")
    child.expect("#")
#login timeout
if result == 0:
    print ("MGMT could not login.")
    sys.exit()
print ("login to MGMT successfully.")
print ("begin to connect release info, please waiting...")

#get the hardware version
child.sendline("show hardware detail | grep Chassis")
time.sleep(5.0)
child.expect("#")
chassInfo = child.before
#print ("hardWare version is %s" % child.before)
#print ("=================================")


#open COM CLI
showResult = ""
showGroup = ""
if os.path.exists(wmgpm):
    print ("This is wmg.")
    child.sendline("st o")
    child.expect(">")
    child.sendline("ManagedElement=1")
    child.expect(">")
    child.sendline("configure")
    child.expect(">")
    child.sendline("WmgFunction=1")
    #print("@@@@@@@@@@@@@@@@@@@@@@@@")
    child.expect(">")
    child.sendline('SystemManagement=1')
    child.expect(">")
    child.sendline('showSystem')
    child.expect(">")
    showResult = child.before
    time.sleep(3.0)
    child.sendline('showGroupState')
    child.expect(">")
    #print (child.before)
    showGroup = child.before

    if os.path.exists(wmgPdcArchive):
        pass
    else:
        os.makedirs(wmgPdcArchive)
    open_file = open(wmgPdcArchive +"config_data.log","w+")
else:
    print ("This is epdg.")
    child.sendline("st o")
    child.expect(">")
    child.sendline("ManagedElement=1")
    child.expect(">")
    child.sendline("configure")
    child.expect(">")
    child.sendline("EpdgFunction=1")
    #print("@@@@@@@@@@@@@@@@@@@@@@@@")
    child.expect(">")
    child.sendline('SystemManagement=1')
    child.expect(">")
    child.sendline('showSystem')
    child.expect(">")
    showResult = child.before
    time.sleep(3.0)
    child.sendline('showGroupState')
    child.expect(">")
    #print (child.before)
    showGroup = child.before

    if os.path.exists(epdgPdcArchive):
        pass
    else:
        os.makedirs(epdgPdcArchive)
    open_file = open(epdgPdcArchive +"config_data.log","w+")

#print (showResult)
#print (showGroup)
hardware = re.compile("(.*\n)+(Chassis Type.*)")
hardwareVersion = hardware.match(chassInfo).group(2)
#print("hardwareVersion = %s" % hardwareVersion)

pdtvsn = re.compile('(.*\n)+(Product Version.*)')
productStr = pdtvsn.match(showResult).group(2)
#print ("productStr = %s" % productStr)

groups = []
#get group info
if showGroup != None:
    print ('++++++++++++++++++++++')
    eachGroup = showGroup.split('-\r\n')[1:-1]
    print ('len(eachGroup) = %s' % len(eachGroup))
    #print ('eachGroup = %s' % eachGroup)
    if len(eachGroup) > 0:
        for item in eachGroup:
            #get each asp
            temp = item.split('\r\n')
            #print ('temp = %s' % temp)
            #asp1 = temp[0].split('\s+')
            asp1 = re.split('\s+', temp[0])
            #print ('asp1 = %s' % asp1)
            if "-----" in temp[1]:
               group = 'Group' + asp1[0] + '=' + asp1[1]
            else:
               asp2 = re.split('\s+', temp[1])
               #print ('asp2 = %s' % asp2)
               group = 'Group' + asp1[0] + '=' + asp1[1] + ', ' + asp2[1].strip()
            #print ('=====%s' % group)
            groups.append(group)

#get detail release
product = (productStr.split(':')[1]).strip()
rpv = product.split('_')
release = ''
if rpv[-1].strip() in releaseList:
    release = "2014B"
else:
    release = "2015B"

#write the infos in file

#open_file = open("/md/wmg/pdc/archive/temp/tmp/config_data.log","w+")
open_file.write("Host Name=%s" % hostname)
open_file.write('\n')
open_file.write("HW=%s" % (hardwareVersion.split(':')[1]).strip())
open_file.write('\n')
open_file.write("Release=%s" % product)
open_file.write('\n')
#open_file.write("Product Version: %s" % product)
#open_file.write('\n')
open_file.write('[Groups]')
open_file.write('\n')
print ('groups = %s' % groups)
for group in groups:
    open_file.write(group)
    open_file.write('\n')
open_file.close()
print ("infos connected successfully.")

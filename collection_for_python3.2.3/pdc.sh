#!/bin/sh
help()
{
    echo "###############################################"
    echo "## Usage:                                    ##"
    echo "##    ./pdc.sh or ./pdc.sh -o                ##"
    echo "## -o on demand connecting all data          ##"
    echo "##                                           ##"
    echo "###############################################"
}

para=$1
echo "para=$para"

if [ "$para" != "" ] && [ "$para" != "-o" ];then
    help
    exit 1
fi

#get the path of current script shell
basepath=$(cd `dirname $0`; pwd)

MAINPATH="/md/wmg/pdc/archive/"
MONTHLYPATH="${MAINPATH}monthly/"
MONTH=`date +%Y%m`
ENMONTH=`date | awk -F ' ' '{print $2}'`
DAY=`date +%d`
ZONE=`date -R | awk -F ' ' '{print $6}'`
HOSTNAME=`hostname`
#echo ${#HOSTNAME}
HOSTNAME_LEN=${#HOSTNAME}
#echo $HOSTNAME_LEN

if [ $HOSTNAME_LEN -gt 18 ];then
   #length of hostname should be less or equal to 18
   HOSTNAME=${HOSTNAME:0:18}
   #If there is a "." In hostname, replace it with "-"
   HOSTNAME=${HOSTNAME//./-}
   #echo $HOSTNAME
fi

#copy pm files
epdgpm=/md/epdg/pm/
wmgpm=/md/wmg/pm/

if [ -d /tmp/pm ];then
    rm -rf $path/pm
fi

if [ -d $wmgpm ];then
    cp -rf $wmgpm /tmp/
    cd /tmp/pm
    find ./ -type f -not -name '*.gz' -delete
    folder=`ls /tmp/pm/`
    if [ "$folder" = "" ];then
        echo "there is no PM files in $wmgpm."
        exit 1
    else 
        gunzip *
    fi
else
    if [ -d $epdgpm ];then
        cp -rf $epdgpm /tmp/
        cd /tmp/pm
        find ./ -type f -not -name '*.gz' -delete
        folder=`ls /tmp/pm/`
        if [ "$folder" = "" ];then
            echo "there is no PM files in $epdgpm."
            exit 1
        else
            gunzip *
        fi
    else
        echo "no directory $epdgpm or $wmgpm."
        exit 1
    fi
fi

TARNAME="${ENMONTH}_${HOSTNAME}_${MONTH}${ZONE}-monthly_epdg_pdc.tar.gz"

if [ $DAY -eq 01 ];then
    #python3 ${PATHOFSCRIPT}wmg_pdc_log.py
    #python3 /md/pdc/pyscript/connect_cli.py
    python3 $basepath/pyscript/wmg_pdc_log.py
    if [ -d $MONTHLYPATH ];then
        cd ${MAINPATH}temp
        tar -czf $MONTHLYPATH$TARNAME *
    else
        mkdir -p $MONTHLYPATH
        cd ${MAINPATH}temp
        tar -czf $MONTHLYPATH$TARNAME *
    fi
else
    python3 $basepath/pyscript/connect_cli.py
    if [ "$para" = "" ];then
        echo "begin to connecting PM data..."
        python3 $basepath/pyscript/wmg_pdc_log.py 
    else 
        echo "begin to connect data on demand..."
        python3 $basepath/pyscript/on_demand.py $para
    fi
fi

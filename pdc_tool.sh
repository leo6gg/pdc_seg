#!/bin/sh

param1=$1
param2=$2

path=`pwd`
globalFile=$path'/config/global.cfg'
versionTmp=`sed '/^version =/!d;s/.*=//' $globalFile`
if [ -z $versionTmp ];then
    echo "Please need to configure epdg version in ${globalFile}."
    exit 0
fi
version=`echo $versionTmp | sed -r 's/$[[:space:]]+|[[:space:]]+$//g'`

help_info(){
    echo "#################################################################"
    echo "##                                                             ##"
    echo "## usage:                                                      ##"
    echo "## ./pdc_tool.sh collect | report                              ##"
    echo "##                                                             ##"
    echo "## collect: is indicate all pdc data will be collected.        ##"
    echo "## report : is indicate the latest pdc data will be collected  ##"
    echo "##          and the pdc graph will be generated.               ##"
    echo "## example:                                                    ##" 
    echo "## ./pdc_tool.sh collect                                       ##" 
    echo "##                                                             ##"
    echo "#################################################################"
}

if [ $# -eq 0 ];then
   help_info
   exit 1
fi


if [ $param1 = "collect" ];then
    #version=`echo $version | sed -r 's/$[[:space:]]+|[[:space:]]+$//g'`
    case $version in 
         "ePDG_14B_CP03"|"epdg_14b_cp03")
         echo "begin to collecting pdc datas of epdg cp03, please waiting......!"
         python $path"/bin/cp03_ep_scripts/wmg_pdc_log.py"
         
         ;;

         "ePDG_14B_CP04")
         echo "begin to collecting PDC datas of epdg cp04, please waiting......"
         python $path"/bin/cp04_scripts/wmg_pdc_log.py"
         ;;

         "WMG_15B")
         echo "begin to collecting PDC datas of WMG 15b, please waiting......"
         python $path"/bin/15b_scripts/wmg_pdc_log.py"
         ;;

         "SEG_5")
         echo "begin to collecting PDC datas of SEG, please waiting......"
         python $path"/bin/seg_scripts/wmg_pdc_log.py"
         ;;
		 
         *)
         help_info
     esac

fi

if [ $param1 = "report" ];then
    case $version in
         "ePDG_14B_CP03")
         echo "begin to collecting the latest pdc data of epdg cp03, please waiting......"
         python $path"/bin/cp03_ep_scripts/wmg_pdc_latest_xls.py"
         echo "the latest data generate successfully."
         echo " "
         echo "begin to generating excel data......"
         #python $path"/bin/cp03_ep_scripts/wmg_pdc_xls.py"
         echo " "
         echo "begin to generating graphs for pdc data of epdg cp03, please waiting......"
         python $path"/bin/cp03_ep_scripts/pdc_graph.py"
         ;;

         "ePDG_14B_CP04")
         echo "begin to collecting the latest pdc data of epdg cp04 or wmg 15b, please waiting......"
         python $path"/bin/cp04_scripts/wmg_pdc_latest_xls.py"
         echo "the latest data generate successfully."
         echo " "
         echo "begin to generating excel data......"
         python $path"/bin/cp04_scripts/wmg_pdc_xls.py"
         echo " "
         echo "begin to generating graphs for pdc data of epdg cp04 , please waiting......"
         python $path"/bin/cp04_scripts/pdc_chart.py"
         ;;

         "WMG_15B")
         echo "begin to collecting the latest pdc data of wmg 15b, please waiting......"
         python $path"/bin/15b_scripts/wmg_pdc_latest_xls.py"
         echo "the latest data generate successfully."
         echo " "
         echo "begin to generating excel data......"
         python $path"/bin/15b_scripts/wmg_pdc_xls.py"
         echo " "
         echo "begin to generating graphs for pdc data of wmg 15B, please waiting......"
         python $path"/bin/15b_scripts/pdc_chart.py"
         ;;

        "SEG_5")
         echo "begin to collecting the latest pdc data of seg, please waiting......"
         #python $path"/bin/seg_scripts/wmg_pdc_latest_xls.py"
         echo "the latest data generate successfully."
         echo " "
         echo "begin to generating excel data......"
         #python $path"/bin/seg_scripts/wmg_pdc_xls.py"
         echo " "
         echo "begin to generating graphs for pdc data of seg, please waiting......"
         python $path"/bin/seg_scripts/pdc_chart.py"
         ;;
		 
         *)
         help_info
     esac

fi
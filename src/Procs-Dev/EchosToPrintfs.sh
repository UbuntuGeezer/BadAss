#!/bin/bash
# EchosToPrintfs.sh - change echo,s to printf,s in shell.
# 1/12/24.	wmk.
#
# Usage. bash  EchosToPrintfs.sh <file> [<path>]
#
#	<file> = shell in which to convert echo,s
#	<path> = (optional) path to <file>;
#		default = *PWD
#
# Entry. *TODAY1 = current date mm/dd/yy.
#
# Dependencies.
#
# Modification History.
# ---------------------
# 11/24/23.	wmk.	original shell.
# 1/12/24.	wmk.	SetToday shell path integrated for BadAss/src/Procs-Dev.
# Notes. 
#
# P1=<file>, [P2=<path>]
P1=$1
P2=$2
if [ -z "$P1" ];then
 echo "EchosToPrintfs <file> [<path>] missing parameter(s) - abandoned."
 exit 1
fi
if [ -z "$P2" ];then
 P2=$PWD
fi
#	Environment vars:
if [ -z "$TODAY" ];then
 . $codebase/BadAss/src/Procs-Dev/SetToday.sh
fi
#procbodyhere
sed -i 's?echo?printf "%s\\n"?g' $P2/$P1
sed -i "/#.*---------/a# $TODAY1.\twmk.\t(automated) echo,s to printf,s throughout" \
 $P2/$P1
#endprocbody
if [ $? -eq 0 ];then
echo "  EchosToPrintfs $P1 $P2 successful."
~/sysprocs/LOGMSG "  EchosToPrintfs $P1 $P2 successful."
else
 echo "  EchosToPrintfs $P1 $P2 failed."
 ~/sysprocs/LOGMSG "  EchosToPrintfs $P1 $P2 failed."
fi
# end EchosToPrintfs.sh

#!/bin/bash
# EchosToPrintfs.sh - change echo,s to printf,s in shell.
# 4/8/24.	wmk.
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
# 3/16/24.	wmk.	code checked for 4.0.x compatibility; -v added to SetToday;
#			 P1 preserved across SetToday call.
# 4/8/24.	wmk.	'. added at end of printf,s insertion comment.
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
 lclfile=$P1
 . $codebase/Procs-Dev/SetToday.sh -v
 P1=$lclfile
fi
#procbodyhere
sed -i 's?echo?printf "%s\\n"?g' $P2/$P1
sed -i "/#.*---------/a# $TODAY1.\twmk.\t(automated) echo,s to printf,s throughout." \
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

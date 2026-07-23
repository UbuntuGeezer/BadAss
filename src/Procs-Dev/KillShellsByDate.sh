#!/bin/bash
# KillShellsByDate.sh - Kill shell by inserting illegal command at start.
#	7/22/26.	wmk.
#
# Usage. bash  KillShellsByDate.sh -h|<path> <sdate>
#
#	-h = only display KillShellsByDate shell help
#	path = path to shell files; only processes files
#		with '.sh' file extension
#	<sdate> = (optional) if specified, all shells whose "modified" date
#		precedes this date will be killed "yyyy-mm-dd"
#
# Entry.  <path>/*.sh
#
# Exit.  <path>*.sh files modified with line 
#
# Modification History.
# ---------------------
# 7/22/26.	wmk.	-h option support.
# 7/22/26.	wmk.	modified for Windows 11; makw > gawk.
# 7/22/26.	wmk.	(automated) UnKillShell to reinstate shell. 
# 7/22/26.	wmk.	(automated) Modification History sorted.
# 1/13/24.	wmk.	SetToday path corrected with *libbase, -v option added; 
# 1/13/24.	wmk.	(automated) printf "%s\n",s to printf,s throughout 
# 1/13/24.	wmk.	(automated) Version 3.0.6 Make old paths removed. 
# 1/13/24.	*binpath	corrected to use *libbase. 
# 11/15/23.	wmk.	*codebase for root to shells. 
# 11/3/23.	wmk.	Version 3.0.0 path updates. 
# 10/11/23.	REPLY	printf "%s\n". 
# 10/11/23.	wmk.	revert to WINGIT_PATH for HP2 system; remove ellipsis from 
# 10/4/23.	wmk.	change to use *gitpath replacing WINGIT_PATH. 
# 9/6/23.	wmk.	paths modified for FLsara86777. 
# 9/2/23.	wmk.	modified for MNcrwg44586; default <before-date> 
# 9/2/23.	verification	added. 
# 9/1/23.	wmk.	original code. 
#
# Notes. This shell is an alternative way of preventing other shells from
# executing when they have gone out-of-date (usually because of referenced
# paths changing or system file folder restructuring).
# This process became necessary for dealing with shells resident on a
# Windows-owned drive since even *sudo cannot use *chmod to change the 'x'
# flag (executable).
#
# P1=-h|<path>, P2=<before-date>
P1=$1
# -h option code
if [ "${P1:0:1}" == "-" ];then
 option=${P1,,}
 if [ "$option" == "-h" ];then
  printf "%s\n" "KillShellsByDate - Kill shell by inserting illegal command at start."
  printf "%s\n" "KillShellsByDate.sh -h|<path> <sdate>"
  printf "%s\n" ""
  printf "%s\n" "  -h = only display KillShellsByDate shell help"
  printf "%s\n" "  <p1> = param1 description"
  printf "%s\n" "  <p2> = param2 description"
  printf "%s\n" ""
  printf "%s\n" "Results:"
  printf "%s\n" ""
  # read -p "Enter ctrl-c to remain in Terminal: "
  exit 0
 else
  printf "%s" "KillShellsByDate.sh -h|<path> <sdate>"
  printf "%s\n" " unrecognized option '$P1' - exiting."
  # read -p "Enter ctrl-c to remain in Terminal: "
  exit 1
 fi		# have -h
fi	# have -
P2=$2
if [ -z "$P1" ];then
 printf "%s" "KillShellsByDate.sh -h|<path> <sdate>"
 printf "%s\n" " missing parameter(s)"
 read -p "Enter ctrl-c to remain in Terminal: "
 exit 1
fi
killpath=$P1
if [ "$killpath" == "./" ];then
 msgsep=
else
 msgsep=/
fi
# get today's date *TODAY.
if [ -z "$TODAY" ];then
 lclP1=$P1
. $sp/SetToday.sh -v
P1=$lclP1
fi
# if P2 unspecified, use today's date.
b4date=$P2
if [ -z "$b4date" ];then
 case $- in
 "*i*")
 printf "%s\n" "  KillShellsByDate <path> [<before-date>] default date is TODAY"
 read -p "   OK to proceed (y/n)? :"
 yn=${REPLY^^}
 if [ "$yn" != "Y" ];then
  printf "%s\n" "KillShellsByDate terminated at user request."
  read -p "Enter ctrl-c to remain in Terminal: "
  exit 0
 fi
 ;;
 *)
 ;;
 esac
 b4date=$TODAY
fi
# create list of *.sh files from P1 path.
pushd ./ > /dev/null
binpath=C:/Users/vncwm/linux/BadAss/src/Procs-Dev
cd $killpath
ls -lh *.sh > $TEMP_PATH/fullshells.txt
# use gawk to pare list down to files meeting date criteria.
gawk -v testdate=$b4date -f $binpath/awkkilldates.txt $TEMP_PATH/fullshells.txt\
 > $TEMP_PATH/datedshells.txt
printf "%s\n" " cat *TEMP_PATH/datedshells.txt for file list..."
read -p "Enter ctrl-c when testing...: "
# now loop on files meeting date criteria.
if test -s $TEMP_PATH/datedshells.txt;then
 dfile=$TEMP_PATH/datedshells.txt
 while read -e;do
  fn=$REPLY
  printf "%s\n" "  processing $killpath$msgsep$fn.."
  $binpath/KillShell.sh $fn $killpath
 done < $dfile
fi		# end non-empty datedlist
popd > /dev/null
$sp/LOGMSG "KillShellsByDate $P1 $P2 complete."
# end KillShellsByDate.sh

#!/bin/bash
# 2024-01-13.	wmk.	(automated) Version 3.0.6 paths eliminated (Lenovo).
# KillShellsByDate.sh - Kill shell by inserting illegal command at start.
#	1/13/24.	wmk.
#
# Usage. bash  KillShellsByDate.sh <path> <date>
#
#	path = path to shell files; only processes files
#		with '.sh' file extension
#	<date> = (optional) if specified, all shells whose "modified" date
#		precedes this date will be killed "yyyy-mm-dd"
#
# Entry.  <path>/*.sh
#
# Exit.  <path>*.sh files modified with line 
#
# Modification History.
# ---------------------
# 01/13/24.	wmk.	(automated) echo,s to printf,s throughout
# 01/13/24.	wmk.	(automated) Version 3.0.6 Make old paths removed.
# 1/13/24.	wmk.	SetToday path corrected with *libbase, -v option added;
#			 *binpath corrected to use *libbase.
# 11/3/23.	wmk.	Version 3.0.0 path updates.
# 11/15/23.	wmk.	*codebase for root to shells.
# Legacy mods.
# 10/11/23.	wmk.	revert to WINGIT_PATH for HP2 system; remove ellipsis from
#			 REPLY printf "%s\n".
# Legacy mods.
# 9/1/23.	wmk.	original code.
# 9/2/23.	wmk.	modified for MNcrwg44586; default <before-date>
#			 verification added.
# 9/6/23.	wmk.	paths modified for FLsara86777.
# 10/4/23.	wmk.	change to use *gitpath replacing WINGIT_PATH.
#
# Notes. This shell is an alternative way of preventing other shells from
# executing when they have gone out-of-date (usually because of referenced
# paths changing or system file folder restructuring).
# This process became necessary for dealing with shells resident on a
# Windows-owned drive since even *sudo cannot use *chmod to change the 'x'
# flag (executable).
#
# P1=<path>, P2=<before-date>
P1=$1
P2=$2
if [ -z "$P1" ];then
 printf "%s\n" "KillPath <path> [<before-date>] missing parameter(s)"
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
. $libbase/src/Procs-Dev/SetToday.sh -v
printf "%s\n" "TODAY is '$TODAY'"
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
binpath=$libbase/src/Procs-Dev
cd $killpath
ls -lh *.sh > $TEMP_PATH/fullshells.txt
# use mawk to pare list down to files meeting date criteria.
mawk -v testdate=$b4date -f $binpath/awkkilldates.txt $TEMP_PATH/fullshells.txt\
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
# end KillShellsByDate.sh

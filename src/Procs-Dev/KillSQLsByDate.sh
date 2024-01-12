#!/bin/bash
echo " ** KillSQLsByDate.sh out-of-date **";exit 1
# 2023-11-26.	wmk.	(automated) Version 3.0.6 paths eliminated (HPPavilion).
# 2023-11-12.	wmk.	(automated) Version 3.0.6 paths eliminated (HPPavilion).
# 2023-11-03   wmk.   (automated) Version 3.0.0 path fixes (HPPavilion).
# KillSQLsByDate.sh - Kill shell by inserting illegal command at start.
#	11/3/23.	wmk.
#
# Usage. bash  KillSQLsByDate.sh <path> [<date>] [<type>]
#
#	path = path to shell files; only processes files
#		with '.sh' file extension
#	<date> = (optional) if specified, all shells whose "modified" date
#		precedes this date will be killed "yyyy-mm-dd"
#	<type> = (optional) psq|sql indicates which file suffix,s will be
#		killed; psq = .psq files, sql=.sql files (default=psq)
#
# Entry.  <path>/*.sh
#
# Exit.  <path>*.sh files modified with line 
# Modification History.
# ---------------------
# 11/26/23.	wmk.	(automated) echo,s to printf,s throughout
# 11/26/23.	wmk.	(automated) Version 3.0.6 Make old paths removed.
# 11/12/23.	wmk.	(automated) Version 3.0.6 Make old paths removed.
# 11/3/23.	wmk.	Version 3.0.0 path updates.
# Legacy mods.
# 9/9/23.	wmk.	original code; adapted from KillShellsByDate.
#			 add <type> processing.
# Legacy mods.
# 9/1/23.	wmk.	original code.
# 9/2/23.	wmk.	modified for MNcrwg44586; default <before-date>
#			 verification added.
# 9/6/23.	wmk.	paths modified for FLsara86777.
#
# Notes. This shell is an alternative way of preventing SQL queries from
# executing when they have gone out-of-date (usually because of referenced
# paths changing or system file folder restructuring).
# This process became necessary for dealing with shells resident on a
# Windows-owned drive since even *sudo cannot use *chmod to change the 'x'
# flag (executable).
#
# P1=<path>, P2=<before-date>, P3=<type>
P1=$1
P2=$2
P3=${3,,}
if [ -z "$P1" ];then
 printf "%s\n" "KillSQLsByDate <path> [<before-date>] [psq|sql|all] missing parameter(s)"
 read -p "Enter ctrl-c to remain in Terminal: "
 exit 1
fi
killpsq=0
killsql=0
if [ -z "$P3" ];then
 killpsq=1
else
 case "$P3" in
  "psq")
   killpsq=1
  ;;
  "sql")
   killsql=1
  ;;
  "all")
   killpsq=1
   killsql=1
  ;;
  *)
   printf "%s\n" "KillSQLsByDate <path> [<before-date>] [psq|sql|all] $P3 unrecognized - abandoned."
   read -p "Enter ctrl-c to remain in Terminal: "
   exit 1
  ;;
 esac
fi
killpath=$P1
if [ "$killpath" == "./" ];then
 msgsep=
else
 msgsep=/
fi
# get today's date *TODAY.
if [ -z "$TODAY" ];then
. $WINGIT_PATH/TerrCode86777/FLsara86777cb/Procs-Dev/SetToday.sh
printf "%s\n" "TODAY is '$TODAY'"
fi
# if P2 unspecified, use today's date.
b4date=$P2
if [ -z "$b4date" ];then
 case $- in
 "*i*")
 printf "%s\n" "  KillSQLsByDate <path> [<before-date>] default date is TODAY"
 read -p "   OK to proceed (y/n)? :"
 yn=${REPLY^^}
 if [ "$yn" != "Y" ];then
  printf "%s\n" "KillSQLsByDate terminated at user request."
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
binpath=$WINGIT_PATH/TerrCode86777/FLsara86777cb/Procs-Dev
cd $killpath
if test -f $TEMP_PATH/fullsqls.txt;then rm $TEMP_PATH/fullsqls.txt;fi
if [ $killpsq -ne 0 ];then
 ls -lh *.psq >> $TEMP_PATH/fullsqls.txt
fi
if [ $killsql -ne 0 ];then
 ls -lh *.sql >> $TEMP_PATH/fullsqls.txt
fi
# use mawk to pare list down to files meeting date criteria.
if ! test -f $TEMP_PATH/fullsqls.txt;then
 printf "%s\n" "KillSQLsByDate no .psq/.sql files to process..."
 read -p "Enter ctrl-c to remain in Terminal: "
 exit 0
fi
mawk -v testdate=$b4date -f $binpath/awkkilldates.txt $TEMP_PATH/fullsqls.txt\
 > $TEMP_PATH/datedsqls.txt
printf "%s\n" " cat *TEMP_PATH/datedshells.txt for file list..."
read -p "Enter ctrl-c when testing...: "
# now loop on files meeting date criteria.
if test -s $TEMP_PATH/datedsqls.txt;then
 dfile=$TEMP_PATH/datedsqls.txt
 while read -e;do
  fn=$REPLY
  printf "%s\n" "  processing $killpath$msgsep$fn..."
  $binpath/KillSQL.sh $fn $killpath
 done < $dfile
else
 printf "%s\n" "KillSQLsByDate no .psq/.sql files meet date criteria..."
 read -p "Enter ctrl-c to remain in Terminal: "
 exit 0
fi		# end non-empty datedlist
popd > /dev/null
# end KillSQLsByDate.sh

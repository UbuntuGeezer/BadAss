#!/bin/bash
echo " ** UnKillGShells.sh out-of-date **";exit 1
echo " ** $P1 out-of-date-exiting **";exit 1
# 2023-09-06   wmk.   (automated) ver2.0 path fixes.
# UnKillGShells.sh - UnKill Project-Geany shells by removing exit at start.
#	9/4/23.	wmk.
#
# Usage. bash  UnKillGShells.sh  <date>
#
#	<date> = (optional) if specified, all shells whose "modified" date
#		precedes this date will be killed "yyyy-mm-dd"; default is *TODAY.
#
# Entry.  *PWD = some ../Projects-Geany folder
#		  */*.sh shell files present
#
# Calls. UnKillMultShells.
#
# Exit.  <path>*.sh files modified with line 
#			"ECHO <shell-name> out-of-date;exit 1"
#
# Modification History.
# ---------------------
# 9/4/23.	wmk.	original code; adapted from MNcrwg44586/KillGShellsByDate.
# Legacy mods.
# 9/1/23.	wmk.	original code.
# 9/2/23.	wmk.	./ prefix added when passing path to KillShells.
# 9/2/23.	wmk.	modified for MNcrwg44586.
#
# Notes. This shell loops on all of the Projects-Geany folders removing the
# exit statement that kills
# all of the *.sh files to prevent them from
# executing when they have gone out-of-date (usually because of referenced
# paths changing or system file folder restructuring).
# This process became necessary for dealing with shells resident on a
# Windows-owned drive since even *sudo cannot use *chmod to change the 'x'
# flag (executable).
#
# Path corrections are also made from previous systems:
# *folderbase= ...Windows/Users/Bill -> .../Windows
# *pathbase= ...Terrtories           -> ...Territories/FL/SARA/86777
# *codebase= ...TerritoriesCB        -> ...TerritoriesCB/FLsara86777
P1=$1
# get today's date *TODAY.
if [ -z "$TODAY" ];then
. $WINGIT_PATH/TerritoriesCB/FLsara86777/Procs-Dev/SetToday.sh
echo "TODAY is '$TODAY'"
fi
# check for empty *P1.
if [ -z "$P1" ];then
 case $- in
 "*i*")
 echo "UnKillGShells [<before-date>] - using *TODAY for default."
 read -p "  OK to coninue (y/n): "
 yn=${REPLY^^}
 if [ "$yn" != "Y" ];then
  echo "UnKillGShells stopped by user."
  read -p "Enter ctrl-c to remain in Terminal: "
  exit 0
 fi
 ;;
 *)
 ;;
 esac
 b4date=$TODAY
fi		# <before-date> unspecified
# if P1 unspecified, use today's date.
b4date=$P1
if [ -z "$b4date" ];then
 b4date=$TODAY
fi
pushd ./ > /dev/null
cd $WINGIT_PATH/TerritoriesCB/FLsara86777/Projects-Geany
ls -lh > $TEMP_PATH/Gprojfiles.txt
# use mawk to pare list down to files meeting Geany folder criteria.
mawk '{if(substr($1,1,1) == "d")print $8;;}' $TEMP_PATH/Gprojfiles.txt \
 > $TEMP_PATH/Gprojpaths.txt
mawk -F "/" '{print $NF}' $TEMP_PATH/Gprojpaths.txt \
 > $TEMP_PATH/Gprojfolders.txt
echo " cat *TEMP_PATH/Gprojfolders.txt for folder list..."
# now loop on files meeting date criteria.
if test -s $TEMP_PATH/Gprojfolders.txt;then
 binpath=$WINGIT_PATH/TerritoriesCB/FLsara86777/Procs-Dev
 file=$TEMP_PATH/Gprojfolders.txt
 #climit=10
 ccount=0
 while read -e;do
  fn=$REPLY
  frstchar=${fn:0:1}
  skip=0
  if [ "$frstchar" == "\$" ];then break;fi
  if [ "$frstchar" == "#" ];then skip=1;fi
  if [ ${#fn} -eq 0 ];then skip=1;fi
  if [ $skip -eq 0 ];then
   echo "  processing folder $PWD/$fn..."
   pushd ./ > /dev/null
   cd $fn
   $binpath/UnKillMultShells.sh ./
   popd > /dev/null
   #if [ $ccount -eq $climit ];then break;fi
   ccount=$((ccount++))
  fi	# skip
 done < $file
fi		# end non-empty datedlist
popd > /dev/null
echo " UnKillGShells - $ccount folders processed."
# end UnKillGShells.sh

#!/bin/bash
echo " ** UnKillMultSQLs.sh out-of-date **";exit 1
# 2023-11-26.	wmk.	(automated) Version 3.0.6 paths eliminated (HPPavilion).
# UnKillMultSQLs.sh - UnKill multiple SQL files removing error at start.
#	11/3/23.	wmk.
#
# Usage. bash  UnKillMultSQLs.sh [<path>] [<type>] [-a]
#
#	<path> = (optional) path to <Make-name> file; default is *PWD
#	<type> = (optional) type(s) of SQL to UnKill; psq|sql|all
#		psq = .psq files, sql = .sql files, all = both .psq and .sql files
#	-a = (optional) unkill regardless of date.
#
# Entry.  [<path>/]<Make-name> exists
#		 line 2 is .."(error out-of-date)"
#
# Exit.  [<path>/]<Make-name> modified with line 
#			"\$(error out-of-date)" removed
#		 # *TODAY   wmk. line added.
#
# Modification History.
# ---------------------
# 11/3/23.	wmk.	Version 3.0.0 path updates.
# 11/3/23.	wmk.	shell name corrected in messages.
# 11/26/23.	wmk.	(automated) echo,s to printf,s throughout
# 11/26/23.	wmk.	(automated) Version 3.0.6 Make old paths removed.
# 11/26/23.	wmk.	-a option added.
# Legacy mods.
# 9/9/23.	wmk.	original code; adapted from UnKillMultMakes.
# Legacy mods.
# 9/2/23.	wmk.	original code; adapted from UnKillMake.
# 9/8/23.	wmk.	path change for FLsara86777.
#
# Notes. This shell "undoes" all KillSQL operations on SQL files in the
# specified <path> by removing the
# '-out-of-date-exiting' and '.exit 1' lines from the beginning of the SQL
# files.
#
# P1=<path>, [P2=psq|sql|all], [P3=-a]
P1=$1
P2=${2,,}
P3=${3,,}
if [ -z "$P1" ];then
 printf "%s\n" "UnKillMultSQLs [<path>] [psq|sql|all] [-a] missing parameter(s) - abandoned."
 read -p "Enter ctrl-c to remain in Terminal: "
 exit 1
fi
nodate=0
if [ ! -z "$P3" ];then
 if [ "$P3" != "-a" ];then
  printf "%s\n" "UnKillMultSQLs [<path>] [psq|sql|all] [-a] unrecognized '$P3' - abandoned."
  read -p "Enter ctrl-c to remain in Terminal: "
  exit 1
 else
  nodate=1
 fi
fi
killsql=$PWD
if [ ! -z "$P1" ];then
 killsql=$P1
fi
printf "%s\n" $killsql
unkillsql=0
unkillpsq=0
if [ -z "$P2" ];then
 unkillpsq=1
elif [ "$P2" == "all" ];then
 unkillsql=1
 unkillpsq=1
elif [ "$P2" == "psq" ];then
 unkillpsq=1
elif [ "$P2" == "sql" ];then
 unkillsql=1
else
 printf "%s\n" "UnKillMultMultSQLs [<path>] [psq|sql|all] unrecognized $P2 - abandoned."
 read -p "Enter ctrl-c to remain in TerminaL: "
 exit 1
fi 
pushd ./ > /dev/null
if [ "$killsql" != "./" ];then
 cd $killsql
fi
printf "%s\n" $PWD
t=$TEMP_PATH
if test -r $t/killedsqls.txt;then rm $t/killedsqls.txt;fi
if [ $unkillpsq -ne 0 ];then
  if [ $nodate -eq 0 ];then
   grep -rle "out-of-date-exiting" --include "*.psq" >> $t/killedsqls.txt
  else
   ls *.psq >> $t/killedsqls.txt
  fi
fi
if [ $unkillsql -ne 0 ];then
  if [ $nodate -eq 0 ];then
   grep -rle "out-of-date-exiting" --include "*.sql" >> $t/killedsqls.txt
  else
   ls *sql >> $t/killedsqls.txt
  fi
fi
if [ $nodate -eq 0 ] && ! test -s $t/killedsqls.txt;then
  printf "%s\n" " no out-of-date files found - UnKillMultSQLs.sh exiting."
  read -p "Enter ctrl-c to remain in Terminal: "
  exit 0
fi
#
if [ 1 -eq 0 ];then
 cat $t/killedsqls.txt
 popd > /dev/null
 read -p "Enter ctrl-c to remain in Terminal: "
 exit 0
fi
# loop on TEMP_PATH/killedsqls.txt.
if [ 1 -eq 0 ];then
 cat $t/killedsqls.txt
 read -p "Enter ctrl-c to interrupt :"
 touch $t/killedsqls.txt	# rewind killedsqls.txt
fi
b=$codebase/Procs-Dev
file=$t/killedsqls.txt
while read -e;do
 fn=$REPLY
 printf "%s\n" "   processing $fn ..."
 if [ "$killsql" != "./" ];then
  $b/UnKillSQL.sh $fn $killsql
 else
  $b/UnKillSQL.sh $fn
 fi
 if [ $? -eq 0 ];then
  printf "%s\n" "UnKillMultSQLs $fn successful."
 else
  printf "%s\n" "UnKillMultSQLs $fn failed."
 fi
done < $file
# end UnKillMultSQLs.sh

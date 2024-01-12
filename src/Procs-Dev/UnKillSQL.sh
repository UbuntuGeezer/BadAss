#!/bin/bash
echo " ** UnKillSQL.sh out-of-date **";exit 1
# UnKillSQL.sh - UnKill SQL by removing illegal command at start.
#	11/28/23.	wmk.
#
# Usage. bash  UnKillSQL.sh <SQL-name>
#
#	<SQL-name> = name of SQL to kill (with .sh extension if applies)
#	<path> = (optional) path to <SQL-name> file; default is *PWD
#
# Entry.  [<path>/]<SQL-name> exists
#	line 2 is .."out-of-date"
#
# Exit.  [<path>/]<SQL-name> modified with line 
#			"\$(error out-of-date)" removed
#	# *TODAY   wmk. line added.
#
# Modification History.
# ---------------------
# 10/14/23.	wmk.	*folderbase, *codebase, *pathbase definitions; SetToday
#			 path to *codebase/Procs-Dev
# 11/26/23.	wmk.	(automated) echo,s to printf,s throughout
# 11/26/23.	wmk.	(automated) Version 3.0.6 Make old paths removed.
# 11/26/23.	wmk.	message fixed when P2=./; add comment in Modification
#			 History.
# 11/28/23.	wmk.	replace spaces with tabs in new lines added.
# Legacy mods.
# 9/9/23.	wmk.	original code; adapted from UnKillMake.
#
# Notes. This shell "undoes" a KillSQL operation by removing the
# -out-of-date-exiting and '.exit 1' lines from the beginning of the SQL file.
#
P1=$1
P2=$2
if [ -z "$P1" ];then
 printf "%s\n" "UnKillSQL <SQL-name> [<path>] missing parameter(s)"
 read -p "Enter ctrl-c to remain in Terminal: "
 exit 1
fi
killSQL=$PWD
if [ ! -z "$P2" ];then
 killSQL=$P2
fi
if [ "$killsql" != "\./" ];then
 printf "%s\n" $killSQL/$P1
else
 printf "%s\n" $killSQL$P1
fi
if ! test -s $killSQL/$P1;then
 printf "%s\n" " UnKillSQL - $P1 is empty or non-existent."
 read -p "Enter ctrl-c to remain in Terminal: "
 exit 0
fi
if [ -z "$b" ];then
 export b=$codebase/Procs-Dev
fi
if [ -z "$TODAY" ];then
 $b/SetToday.sh
fi
printf "%s\n" "/out-of-date-exiting/d" \
  > $TEMP_PATH/sedunkillsql.txt
printf "%s\n" "/\.exit 1/d" \
  >> $TEMP_PATH/sedunkillsql.txt
printf "%s\n" "/.*pathbase =.*Territories$/s?Territories?Territories/FL/SARA/86777?1" \
  >> $TEMP_PATH/sedunkillmake.txt
printf "%s\n" "1a-- $TODAY\twmk.\t\(automated\) Version 3.0.6 SQL fixes." >> $TEMP_PATH/sedunkillsql.txt
printf "%s\n" "/-- \* -----/a-- * $TODAY1.\twmk.\t\(automated\) Version 3.0.6 SQL fixes." >> $TEMP_PATH/sedunkillsql.txt
if [ "$killSQL" == "./" ];then
 sed -i -f $TEMP_PATH/sedunkillsql.txt $killSQL$P1
else
 sed -i -f $TEMP_PATH/sedunkillsql.txt $killSQL/$P1
fi
if [ $? -eq 0 ];then
 printf "%s\n" "UnKillSQL $P1 $P2 successful."
else
 printf "%s\n" "UnKillSQL $P1 $P2 failed."
fi
# end UnKillSQL.sh

#!/bin/bash
echo " ** KillSQL.sh out-of-date **";exit 1
# 2023-11-12.	wmk.	(automated) Version 3.0.6 paths eliminated (HPPavilion).
# 2023-11-03   wmk.   (automated) Version 3.0.0 path fixes (HPPavilion).
# KillSQL.sh - Kill SQL query by inserting .exit 1 command at start.
#	11/3/23.	wmk.
#
# Usage. bash  KillSQL.sh <SQL-name> [<path>]
#
#	<shell-name> = name of SQL query to kill (with .psq/sql extension if applies)
#	<path> = (optional) path to <SQL-name> file; default is *PWD
#
# Entry.  [<path>/]<SQL-name> exists 
#
# Exit.  [<path>/]<SQL-name> modified with line 
#			exit 1  -- out-of-date
#
# Modification History.
# ---------------------
# 11/12/23.	wmk.	(automated) Version 3.0.6 Make old paths removed.
# 11/3/23.	wmk.	Version 3.0.0 path updates.
# Legacy mods.
# 9/9/23.	wmk.	original code; adpapted from KillShell.
# Legacy mods.
# 9/1/23.	wmk.	original code.
# 9/2/23.	wmk.	error message texts corrected.
# 9/6/23.	wmk.	code checked for FLsara86777.
# 9/9/23.	wmk.	bug fix determining path separator.
#
# Notes. This shell is an alternative way of preventing a shell from executing
# when it has gone out-of-date (usually because of referenced paths changing).
# It became necessary for dealing with shells resident on a Windows-owned drive
# since even *sudo cannot use *chmod to change the 'x' flag (executable).
#
P1=$1
P2=$2
if [ -z "$P1" ];then
 echo "KillSQL <SQL-name> [<path>] missing parameter(s)"
 read -p "Enter ctrl-c to remain in Terminal: "
 exit 1
fi
killpath=$PWD
if [ ! -z "$P2" ];then
 killpath=$P2
fi
klen=${#killpath}
if [ "${killpath:klen-1:1}" == "/" ];then
 ksep=
else
 ksep=/
fi
echo $killpath$ksep$P1
pushd ./ > /dev/null
cd $killpath
if ! test -s $P1;then
 echo " KillSQL - $P1 is empty or non-existent."
 read -p "Enter ctrl-c to remain in Terminal: "
 exit 0
fi
printf "%s\n%s\n" "1a. \\\n$P1-out-of-date-exiting\\\n" "1a.exit 1     -- $P1 out of-date"  > $TEMP_PATH/sedkillsql.txt
sed -i -f $TEMP_PATH/sedkillsql.txt $P1
if [ $? -eq 0 ];then
 echo "KillSQL $P1 $P2 successful."
else
 echo "KillSQL $P1 $P2 failed."
fi
popd > /dev/null
# end KillSQL.sh

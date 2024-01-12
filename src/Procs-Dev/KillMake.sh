#!/bin/bash
echo " ** KillMake.sh out-of-date **";exit 1
# 2023-11-12.	wmk.	(automated) Version 3.0.6 paths eliminated (HPPavilion).
# KillMake.sh - Kill shell by inserting illegal command at start.
#	11/8/23.	wmk.
#
# Usage. bash  KillMake.sh <shell-name> [<path>]
#
#	<shell-name> = name of shell to kill (with .sh extension if applies)
#	<path> = (optional) path to <shell-name> file; default is *PWD
#
# Entry.  [<path>/]<shell-name> exists 
#
# Exit.  [<path>/]<shell-name> modified with line 
#			"ECHO <shell-name> out-of-date"
#
# Modification History.
# ---------------------
# 11/12/23.	wmk.	(automated) Version 3.0.6 Make old paths removed.
# 11/3/23.	wmk.	Version 3.0.0 path updates.
# 11/8/23.	wmk.	Version 3.0.6 *folderbase, *codebase, *pathbase assumed
#			 set on entry.
# Legacy mods.
# 9/2/23.	wmk.	original code; adapted from KillShell.
# 9/6/23.	wmk.	code checked for FLsara86777.
#
# Notes. This shell is an alternative way of disabling a *make file
# when it has gone out-of-date (usually because of referenced paths changing).
#
P1=$1
P2=$2
if [ -z "$P1" ];then
 echo "KillMake <shell-name> [<path>] missing parameter(s)"
 read -p "Enter ctrl-c to remain in Terminal: "
 exit 1
fi
killpath=$PWD
if [ ! -z "$P2" ];then
 killpath=$P2
fi
if [ "$killpath" == "./" ];then
 echo $P1
else
 echo $killpath/$P1
fi
pushd ./ > /dev/null
cd $killpath
if ! test -s $P1;then
 echo " KillMake - $P1 is empty or non-existent."
 read -p "Enter ctrl-c to remain in Terminal: "
 exit 0
fi
printf "%s\n%s\n" "1a""\$(error out-of-date)" "" > $TEMP_PATH/sedkillmake.txt
sed -i -f $TEMP_PATH/sedkillmake.txt $P1
if [ $? -eq 0 ];then
 echo "KillMake $P1 $P2 successful."
else
 echo "KillMake $P1 $P2 failed."
fi
popd > /dev/null
# end KillMake.sh

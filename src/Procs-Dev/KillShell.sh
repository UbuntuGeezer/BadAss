#!/bin/bash
echo " ** KillShell.sh out-of-date **";exit 1
# KillShell.sh - Kill shell by inserting illegal command at start.
#	11/15/23.	wmk.
#
# Usage. bash  KillShell.sh <shell-name> [<path>]
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
# 11/8/23.	wmk.	*folderbase, *pathbase, *codebase assumed on entry.
# 11/12/23.	wmk.	(automated) Version 3.0.6 Make old paths removed.
# 11/15/23.	wmk.	verified for Lenovo system.
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
# P1=<shell-name>, [P2=<path>]
P1=$1
P2=$2
if [ -z "$P1" ];then
 echo "KillShell <shell-name> [<path>] missing parameter(s)"
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
 echo " KillShell - $P1 is empty or non-existent."
 read -p "Enter ctrl-c to remain in Terminal: "
 exit 0
fi
sed -i -f $TEMP_PATH/sedkillsh.txt $P1
if [ $? -eq 0 ];then
 echo "KillShell $P1 $P2 successful."
else
 echo "KillShell $P1 $P2 failed."
fi
popd > /dev/null
# end KillShell.sh

#!/bin/bash
# 2024-01-14.	wmk.	(automated) UnKillShell to reinstate shell (Lenovo).
# UnKillMultShells.sh - UnKill multiple .sh files removing exit at start.
#	1/14/24.	wmk.
#
# Usage. bash  UnKillMultShells.sh <path> [-a]
#
#	<path> = path to <Make-name> file
#	-a	= (optional) ALL - run UnKill on all .sh files prior to date
#		  regardless of whether they have the killing line present.
#		 (use to force folder path corrections)
#
# Entry.  <path>/.sh files exist with
#		 line 2  .."out-of-date.."
#
# Exit.  [<path>/.sh file modified with line 
#			"..error out-of-date.." removed
#		 # *TODAY   wmk. line added.
#
# Modification History.
# ---------------------
# 1/14/24.	wmk.	(automated) echo,s to printf,s throughout
# 1/14/24.	wmk.	(automated) UnKillShell to reinstate shell.
# 1/14/24.	wmk.	*b path changed to use *codebase.
# 11/12/23.	wmk.	(automated) Version 3.0.6 Make old paths removed.
# 11/3/23.	wmk.	Version 3.0.0 path updates.
# 11/3/23.	wmk.	replace missing *grep at line 96.
#
# Notes. This Make "undoes" all KillShell operations on *.sh files in the
# specified <path> by removing the
# \$(error out-of-date) line from the beginning of the *make file
# and adjusting the folderbase = Make to match ver2.0.
#
# P1=<path>, P2=-a
P1=$1
P2=${2,,}
if [ -z "$P1" ];then
 printf "%s\n" "UnKillMultShells <path> [-a] missing parameter(s)"
 read -p "Enter ctrl-c to remain in Terminal: "
 exit 1
fi
doall=0
if [ ! -z "$P2" ];then
 if [ "$P2" == "-a" ];then
  doall=1
 else
  printf "%s\n" "UnKillMultShells <path> [-a] unrecognized option $P2 - abandoned."
  read -p "Enter ctrl-c to remain in Terminal: "
  exit 1
 fi
fi
killShell=$PWD
if [ ! -z "$P1" ];then
 killShell=$P1
fi
printf "%s\n" $killShell
pushd ./ > /dev/null
if [ "$killShell" != "./" ];then
 cd $killShell
fi
printf "%s\n" $PWD
t=$TEMP_PATH
if [ $doall -eq 0 ];then
 grep -rle "printf "%s\n".*out-of-date.*exit" > $TEMP_PATH/killedshells.txt
 if [ $? -ne 0 ];then
  printf "%s\n" " no out-of-date files found - UnKillMultShells.sh exiting."
  read -p "Enter ctrl-c to remain in Terminal: "
  exit 0
 fi
else
 ls *.sh > $t/killedshells.txt
fi
printf "%s\n" "  back from grep..."
if [ 1 -eq 0 ];then
 cat $t/killedshells.txt
 popd > /dev/null
 read -p "Enter ctrl-c to remain in Terminal: "
 exit 0
fi
# loop on TEMP_PATH/killedshells.txt.
if [ 1 -eq 0 ];then
 cat $t/killedshells.txt
 read -p "Enter ctrl-c to interrupt :"
 touch $t/killedshells.txt	# rewind killedshells.txt
fi
b=$codebase/src/Procs-Dev
file=$t/killedshells.txt
while read -e;do
 fn=$REPLY
 printf "%s\n" "   processing $fn ..."
 $b/UnKillShell.sh $fn
 if [ $? -eq 0 ];then
  printf "%s\n" "UnKillShell $fn successful."
 else
  printf "%s\n" "UnKillShell $fn failed."
 fi
done < $file
# end UnKillMultShells.sh

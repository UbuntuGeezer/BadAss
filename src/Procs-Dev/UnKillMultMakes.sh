#!/bin/bash
# 2024-01-14.	wmk.	(automated) UnKillShell to reinstate shell (Lenovo).
# UnKillMakes.sh - UnKill multiple *make files removing error at start.
#	1/14/23.	wmk.
#
# Usage. bash  UnKillMultMakes.sh <path> [-a]
#
#	<path> = path to <Make-name> file; default is *PWD
#	-a = (optional) if specified, do ALL makefiles regardless
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
# 1/14/24.	wmk.	(automated) echo,s to printf,s throughout
# 1/14/24.	wmk.	(automated) UnKillShell to reinstate shell.
# 1/14/24.	wmk.	*b changed to use *codebase.
# 11/12/23.	wmk.	(automated) Version 3.0.6 Make old paths removed.
# 11/3/23.	wmk.	Version 3.0.0 path updates.
# Legacy mods.
# 9/2/23.	wmk.	original code; adapted from UnKillMake.
# 9/8/23.	wmk.	path change for FLsara86777.
# 10/11/23.	wmk.	"UnKillMultMakes" in messages; ellipsis removed from
#			 processing message; -a parameter support, pushd, popd added.
#
# Notes. This Make "undoes" all KillMake operations on *make files in the
# specified <path> by removing the
# \$(error out-of-date) line from the beginning of the *make file
# and adjusting the folderbase = Make to match ver2.0.
#
# P1=<path>, [P2=-a]
P1=$1
P2=${2,,}
if [ -z "$P1" ];then
 printf "%s\n" "UnKillMultMakes <path> [-a] missing parameter(s) - abandoned."
 read -p "Enter ctrl-c to remain in Terminal: "
 exit 1
fi
isall=0
if [ ! -z "$P2" ];then
 if [ "$P2" != "-a" ];then
  printf "%s\n" "UnKillMultMakes <path> [-a] unrecognized '$P2' - abandoned."
  read -p "Enter ctrl-c to remain in Terminal: "
  exit 1
 else
  isall=1
 fi
fi
killMake=$PWD
if [ ! -z "$P1" ];then
 killMake=$P1
fi
printf "%s\n" $killMake
#procbodyhere
pushd ./ > /dev/null
if [ "$killmake" != "./" ];then
 cd $killMake
fi
printf "%s\n" $PWD
t=$TEMP_PATH
if [ $isall -eq 0 ];then
 grep -rle "error.*out-of-date" --include "Make*" > $t/killedmakes.txt
 printf "%s\n" "  back from grep..."
 if [ $? -ne 0 ];then
  printf "%s\n" " no out-of-date files found - UnKillMultMakes.sh exiting."
  read -p "Enter ctrl-c to remain in Terminal: "
  exit 0
 fi
 if [ 1 -eq 0 ];then
  cat $t/killedmakes.txt
  popd > /dev/null
  read -p "Enter ctrl-c to remain in Terminal: "
  exit 0
 fi
else
 ls Make* > $t/killedmakes.txt
fi	# end isall
# loop on TEMP_PATH/killedmakes.txt.
if [ 1 -eq 0 ];then
 cat $t/killedmakes.txt
 read -p "Enter ctrl-c to interrupt :"
 touch $t/killedmakes.txt	# rewind killedmakes.txt
fi
b=$codebase/src/Procs-Dev
file=$t/killedmakes.txt
while read -e;do
 fn=$REPLY
 printf "%s\n" "   processing $fn .."
 $b/UnKillMake.sh $fn
 if [ $? -eq 0 ];then
  printf "%s\n" "UnKillMake $fn successful."
 else
  printf "%s\n" "UnKillMake $fn failed."
 fi
done < $file
popd > /dev/null
#endprocbody
# end UnKillMultMakes.sh

#!/bin/bash
# UnKillShell.sh - UnKill shell by removing illegal command at start.
#	1/12/24.	wmk.
#
# Usage. bash  UnKillShell.sh <shell-name> [<path>]
#
#	<shell-name> = name of shell to kill (with .sh extension if applies)
#	<path> = (optional) path to <shell-name> file; default is *PWD
#
# Entry.  [<path>/]<shell-name> exists
#		 line 2 is .."out-of-date"
#
# Exit.  [<path>/]<shell-name> modified with line 
#			"ECHO <shell-name> out-of-date" removed
#		 # *TODAY   wmk. line added.
#
# Modification History.
# ---------------------
# 11/3/23.	wmk.	Version 3.0.0 path updates.
# 11/5/23.	wmk.	change to use tab within inserted lines; bug fix 76777 to
#			 86777.
# 11/8/23.	wmk.	Version 3.0.6 path updates eliminating *folderbase,
#			 *pathbase, *codebase defnitions from shell.
# 11/9/23.	wmk.	TerritoriesCB/FLsara86777 > TerrCode86777/FLsara86777cb
#			 *sed directive added; Modification History > comment added; add
#			 *sed directive to fix 76777 anomalies in SetToday path.
# 11/11/23.	wmk.	*sed to remove *folderbase, *pathbase, *codebase definitions
#			 from shell (11/8 mod lost).
# 11/24/23.	wmk.	run EchosToPrintfs shell on all shells.
# 1/12/24.	wmk.	*thisproj corrected for BadAss/src/Procs-Dev.
# Legacy mods.
# 10/11/23.	wmk.	revert to *WINGIT_PATH for HP2 system.
# Legacy mods.
# 9/1/23.	wmk.	original code.
# 9/2/23.	wmk.	modified for MNcrwg44586.
# 9/6/23.	wmk.	code checked for FLsara86777; printf,s adjusted.
# 9/8/23.	wmk.	ver2.0 fix; correct ~/GitHub paths.
# 9/10/23.	wmk.	error message corrected; add *TEMP_PATH/scratchfile to
#			 changes made.
# 10/5/23.	wmk.	HP-Pavilion; change WINGIT_PATH to git_path in shells.
#
# Notes. This shell restores a shell from execution after being "killed" by
# KillShell. This shellbecame necessary for dealing with shells resident on
# a Windows-owned drive since even *sudo cannot use *chmod to change the '+-x'
# flag (executable).
#
# P1=<shell-name>, [P2=<path>]
P1=$1
P2=$2
if [ -z "$P1" ];then
 echo "UnKillPath <shell-name> [<path>] missing parameter(s)"
 read -p "Enter ctrl-c to remain in Terminal: "
 exit 1
fi
killpath=$PWD
if [ ! -z "$P2" ];then
 killpath=$P2
fi
if [ "$killpath" == "./" ];then
 msgsep=
else
 msgsep=/
fi
echo $killpath$msgsep$P1
if ! test -s $killpath$msgsep$P1;then
 echo " UnKillShell - $P1 is empty or non-existent."
 read -p "Enter ctrl-c to remain in Terminal: "
 exit 0
fi
thispath=$codebase/BadAss/src/Procs-Dev
if [ 1 -eq 0 ];then
# ----- old block editing paths -----
printf "%s\n" "/terrbase=.*\/Users\/Bill/s?\/Users\/Bill??1" > $TEMP_PATH/sedunkillsh.txt
printf "%s\n" "/~\/GitHub/s?~/GitHub?\$git_path?1" >> $TEMP_PATH/sedunkillsh.txt
printf "%s\n" "/pathbase=.*\/Territories$/s?/Territories?/Territories/FL/SARA/86777/?1" >> $TEMP_PATH/sedunkillsh.txt
printf "%s\n" "/pathbase=.*\/TerritoriesCB$/s?/TerritoriesCB/FLsara86777?/TerrCode86777/FLsara86777cb?1" >> $TEMP_PATH/sedunkillsh.txt
printf "%s\n" "/pathbase=.*\/TerritoriesCB/s?pathbase=?codebase=?1" >> $TEMP_PATH/sedunkillsh.txt
printf "%s\n" "/codebase=.*\/TerritoriesCB\//s?/TerritoriesCB/FLsara86777?/TerrCode86777/FLsara86777cb?1" >> $TEMP_PATH/sedunkillsh.txt
printf "%s\n" "/TerritoriesCB\/Procs-Dev/s?/TerritoriesCB/FLsara86777?/TerrCode86777/FLsara86777cb?1" >> $TEMP_PATH/sedunkillsh.txt
printf "%s\n" "/\$TEMP_PATH\/scratchfile/s?\$TEMP_PATH/scratchfile?/dev/null?1" >> $TEMP_PATH/sedunkillsh.txt
#printf "%s\n" "/.*=.*WINGIT_PATH/s?WINGIT_PATH?git_path?1" >> $TEMP_PATH/sedunkillsh.txt
#printf "%s\n" "/.*WINGIT_PATH\//s?WINGIT_PATH/?git_path/?1" >> $TEMP_PATH/sedunkillsh.txt
printf "%s\n" "/.*crmgit_path\//s?crmgit_path/?WINGIT_PATH/?1" >> $TEMP_PATH/sedunkillsh.txt
printf "%s\n" "/folderbase=.*\/Users\/Bill/s?\/Users\/Bill??1"  >> $TEMP_PATH/sedunkillsh.txt
# ----- end old block editing paths -----
fi	# end 1=0
#
printf "%s\n" "/out-of-date.*;.*exit/d" > $TEMP_PATH/sedunkillsh.txt
printf "%s\n" "/if.*ubuntu\"/,/fi/d" >> $TEMP_PATH/sedunkillsh.txt
if [ 1 -eq 0 ];then # below items are now in sedkillpathdefs.txt
 printf "%s\n" "/if.*folderbase\"/,/fi/d" >> $TEMP_PATH/sedunkillsh.txt
 printf "%s\n" "/if.*pathbase\"/,/fi/d" >> $TEMP_PATH/sedunkillsh.txt
 printf "%s\n" "/if.*codebase\"/,/fi/d" >> $TEMP_PATH/sedunkillsh.txt
 printf "%s\n" "/if.*conglib\"/,/fi/d" >> $TEMP_PATH/sedunkillsh.txt
 printf "%s\n" "/if.*congterr\"/,/fi/d" >> $TEMP_PATH/sedunkillsh.txt
fi
printf "%s\n%s\n" "/# -----/a# $TODAY1.\twmk.\t\(automated\) Version 3.0.6 Make old paths removed." >> $TEMP_PATH/sedunkillsh.txt
printf "%s\n" "s?TerritoriesCB/FLsara86777?TerrCode86777/FLsara86777cb?g" \
  >> $TEMP_PATH/sedunkillsh.txt
printf "%s\n" "s?TerritoriesCB/FLsara76777?TerrCode86777/FLsara86777cb?g" \
  >> $TEMP_PATH/sedunkillsh.txt
printf "%s\n" "1a# $TODAY.\twmk.\t(automated) Version 3.0.6 paths eliminated (HPPavilion)." >> $TEMP_PATH/sedunkillsh.txt
if [ "$killpath" == "./" ];then
 # sed -i -f $TEMP_PATH/sedunkillsh.txt $killpath$P1
 sed -i -f $thispath/sedkillpathdefs.txt $killpath$P1
 sed -i -f $TEMP_PATH/sedunkillsh.txt $killpath$P1
else
 # sed -i -f $TEMP_PATH/sedunkillsh.txt $killpath/$P1
 sed -i -f $thispath/sedkillpathdefs.txt $killpath/$P1
 sed -i -f $TEMP_PATH/sedunkillsh.txt $killpath/$P1
fi
if [ $? -eq 0 ];then
 echo "UnKillShell $P1 $P2 successful."
else
 echo "UnKillShell $P1 $P2 failed."
fi
$thispath/EchosToPrintfs.sh $P1 $P2
# end UnKillShell.sh

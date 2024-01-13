#!/bin/bash
# 2024-01-13.	wmk.	(automated) Version 3.0.6 paths eliminated (Lenovo).
# UnKillMake.sh - UnKill Make by removing illegal command at start.
#	1/13/24.	wmk.
#
# Usage. bash  UnKillMake.sh <Make-name> [<Make>]
#
#	<Make-name> = name of Make to kill (with .sh extension if applies)
#	<path> = (optional) path to <Make-name> file; default is *PWD
#
# Entry.  [<path>/]<Make-name> exists
#		 line 2 is .."out-of-date"
#
# Exit.  [<path>/]<Make-name> modified with line 
#			"\$(error out-of-date)" removed
#		 # *TODAY   wmk. line added.
#
# Modification History.
# ---------------------
# 01/13/24.	wmk.	(automated) echo,s to printf,s throughout
# 01/13/24.	wmk.	(automated) Version 3.0.6 Make old paths removed.
# 1/13/24.	wmk.	SetToday corrected; *libbase replaces codebase; /src added
#			 to *libbase paths; printf changed to 'UnKillMake to reinstate
#			 makefile.'
# 11/5/23.	wmk.	Version 3.0.0 path editing.
# 11/8/23.	wmk.	Version 3.0.6 eliminate old *folderbase, *pathbase,
#			 *codebase path definitions from makefile.
# 11/9/23.	wmk.	add Modification History comment to *sed directives; change
#			 match pattern for *folderbase endif to ^endif.
# 11/19/23.	wmk.	Lenovo added to printf header comment.
# Legacy mods.
# 10/5/23.	wmk.	ver2.0 paths checked.
# 10/5/23.	wmk.	improve 'pathbase=' changes by allowing 0..n spaces between
#			 pathbase and '='.
# 9/2/23.	wmk.	original code; adapted from UnKillShell.
# 9/8/23.	wmk.	change from MNcrwg44586 to FLsara86777.
# 9/11/23.	wmk.	eliminate ' ' when fixing *codebase..TerritoriesCB;
#			 change (pathbase)/include, /Projects-Geany, /Procs-Dev to use
#			 (codebase); (codepath) -> (codebase); (basepath)/Projects-Geany
#			 -> (codebase)/Projects-Geany.
#
# Notes. This Make "undoes" a KillMake operation by removing the
# \$(error out-of-date) line from the beginning of the *make file
# and adjusting the folderbase = Make to match ver2.0.
#
P1=$1
P2=$2
if [ -z "$P1" ];then
 printf "%s\n" "UnKillMake <Make-name> [<path>] missing parameter(s)"
 read -p "Enter ctrl-c to remain in Terminal: "
 exit 1
fi
killMake=$PWD
if [ ! -z "$P2" ];then
 killMake=$P2
fi
printf "%s\n" $killMake/$P1
if ! test -s $killMake/$P1;then
 printf "%s\n" " UnKillMake - $P1 is empty or non-existent."
 read -p "Enter ctrl-c to remain in Terminal: "
 exit 0
fi
if [ -z "$b" ];then
 export b=$libbase/src/Procs-Dev
fi
if [ -z "$TODAY" ];then
 . $b/SetToday.sh -v
fi
printf "%s\n" "/ndef.*folderbase/,/^endif/d"  > $TEMP_PATH/sedunkillmake.txt
printf "%s\n" "/ifeq.*\$USER/,/endif/d"  >> $TEMP_PATH/sedunkillmake.txt
printf "%s\n" "/ifndef.*pathbase/,/endif/d"  >> $TEMP_PATH/sedunkillmake.txt
printf "%s\n" "/ifndef.*codebase/,/endif/d"  >> $TEMP_PATH/sedunkillmake.txt
printf "%s\n%s\n" "/# -----/a# $TODAY1.\twmk.\t\(automated\) UnKill to reinstate makefile." \
  >> $TEMP_PATH/sedunkillmake.txt
# 01/13/24.	wmk.	(automated) Version 3.0.6 Make old paths removed.

printf "%s\n" "/(pathbase)\/include/s?(pathbase)?(codebase)?1" \
  >> $TEMP_PATH/sedunkillmake.txt
printf "%s\n" "/(pathbase)\/Projects-Geany/s?(pathbase)?(codebase)?1" \
  >> $TEMP_PATH/sedunkillmake.txt
printf "%s\n" "/(basepath)\/Projects-Geany/s?(basepath)?(codebase)?1" \
  >> $TEMP_PATH/sedunkillmake.txt
printf "%s\n" "/(pathbase)\/Procs-Dev/s?(pathbase)?(codebase)?1" \
  >> $TEMP_PATH/sedunkillmake.txt
printf "%s\n" "/(codepath)/s?(codepath)?(codebase)?1" \
  >> $TEMP_PATH/sedunkillmake.txt
printf "%s\n%s\n" "/.*(error out-of-date.*)/d"  "1a# $TODAY.\twmk.\t\(automated\) UnKill (Lenovo) to reinstate makefile." \
  >> $TEMP_PATH/sedunkillmake.txt
if [ "$killMake" == "./" ];then
 sed -i -f $TEMP_PATH/sedunkillmake.txt $killMake$P1

else
 sed -i -f $TEMP_PATH/sedunkillmake.txt $killMake/$P1
fi
if [ $? -eq 0 ];then
 printf "%s\n" "UnKillMake $P1 $P2 successful."
else
 printf "%s\n" "UnKillMake $P1 $P2 failed."
fi
# end UnKillMake.sh

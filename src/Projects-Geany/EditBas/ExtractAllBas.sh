#!/bin/bash
# 2023-11-13.	wmk.	(automated) Version 3.0.6 paths eliminated (HPPavilion).
# 2023-11-13.	wmk.	(automated) Version 3.0.9 *libpath introduced (HPPavilion).
# ExtractAllBas.sh - Extract all FLsara86777 .bas blocks from .xba module.
# 8/23/23.	wmk.
#
# Usage. bash  ExtractAllBas.sh <xbamodule>
#
#	<xbamodule> = .xba module name (in folder /Basic)
#
# Entry. Basic/<xbamodule>Bas.txt = list of modules in <xbamodule>.xba
#
# Dependencies.
#
# Exit.	/Basic has all .bas blocks extracted from <xbamodule>.xba
#	into files <blockname>.bas
#
# Modification History.
# ---------------------
# 8/22/23.	wmk.	*pathbase, *codebase unconditional.
# 8/23/23.	wmk.	ver2.0 mods to use /src as parent folder.
# 11/13/23.	wmk.	(automated) Version 3.0.6 Make old paths removed.
# 11/13/23.	wmk.	(automated) Version 3.0.9 *libpath introduced.
# Legacy mods.	
# 6/24/23.	wmk.	edited for FLsara86777 from Territories.
# Legacy mods.
# 6/19/23.	wmk.	original shell.
#
# Notes. 
#
# P1=<xbamodule>
#
P1=$1
if [ -z "$P1" ];then
 echo "ExtractAllBas <xbamodule> missing parameter(s) - abandoned."
 read -p "Enter ctrl-c to remain in Terminal: "
 exit 1
fi
#	Environment vars:
#if [ -z "$TODAY" ];then
# . ~/GitHub/TerritoriesCB/Procs-Dev/SetToday.sh
#TODAY=2022-04-22
#fi
#procbodyhere
projpath=$libbase/src/Projects-Geany/EditBas
shellpath=$libbase/src/Procs-Dev
listsuffx=Bas.txt
# loop on ExtractBas.sh (DoSed, make -f MakeExtractBas) with
# list from Basic/$P1Bas.txt file
#$shellpath/LoopAnyShell.sh $projpath/ExtractBas.sh \
# $pathbase/Basic/$P1$listsuffx
pushd ./ > /dev/null
#
echo "  LoopAnyShell beginning processing."
error_counter=0		# set error counter to 0
IFS="&"			# set & as the word delimiter for read.
file=$libbase/src/Basic/$P1/$P1$listsuffx
i=0
while read -e; do
  #reading each line
  echo -e " processing $REPLY " >> $TEMP_PATH/scratchfile
  len=${#REPLY}
  len1=$((len-1))
  firstchar=${REPLY:0:1}
  next_one=$REPLY
#  echo -e "  $firstchar\n is first char of line." >> $HOME/temp/scratchfile
  #expr index $string $substring
  if [ "$firstchar" == "#" ]; then			# skip comment
   echo $REPLY >> $TEMP_PATH/scratchfile
  elif [ "$firstchar" == "\$" ];then
   break
  else
    $projpath/ExtractBas.sh $P1 $next_one
  fi
  i=$((i+1))
done < $file
echo " $i $P2 lines processed."
popd > /dev/null
#endprocbody
echo "  ExtractAllBas $P1 complete."
~/sysprocs/LOGMSG "  ExtractAllBas $P1 complete."
# end ExtractAllBas.sh

#!/bin/bash
# ExtractAllBas.sh - Extract all BadAss .bas blocks from .xba module.
#	4/19/25.	wmk.
#
# Usage. bash  ExtractAllBas.sh -h|<xbamodule>
#
#	-h = only display ExtractAllBas shell help
#	<xbamodule> = .xba module name (in folder /Basic)
#
# Entry. Basic/<xbamodule>Bas.txt = list of modules in <xbamodule>.xba
#
# Exit.	/Basic has all .bas blocks extracted from <xbamodule>.xba
#	into files <blockname>.bas
#
# Modification History.
# ---------------------
# 4/19/25.	wmk.	(automated) Modification History sorted.
# 4/19/25.	wmk.	-h option support. 
# 4/19/25.	wmk.	(automated) Modification History sorted. 
# 1/13/24.	wmk.	code checked for BadAss/src. 
# 11/13/23.	wmk.	(automated) Version 3.0.9 *libpath introduced. 
# 11/13/23.	wmk.	(automated) UnKillShell to reinstate shell. 
# 8/23/23.	wmk.	ver2.0 mods to use /src as parent folder. 
# 8/22/23.	wmk.	*pathbase, *codebase unconditional. 
# 6/24/23.	wmk.	edited for FLsara86777 from Territories. 
# 6/19/23.	wmk.	original shell.
#
# P1=-h|<xbamodule>
#
P1=$1
# -h option code
if [ "${P1:0:1}" == "-" ];then
 option=${P1,,}
 if [ "$option" == "-h" ];then
  printf "%s\n" "ExtractAllBas - Extract all BadAss .bas blocks from .xba module."
  printf "%s\n" "ExtractAllBas.sh -h|<xbamodule>"
  printf "%s\n" ""
  printf "%s\n" "  -h = only display ExtractAllBas shell help"
  printf "%s\n" "  <xbamodule> = .xba module name (in folder /Basic)"
  printf "%s\n" ""
  printf "%s\n" "Results: /Basic has all .bas blocks extracted from <xbamodule>.xba"
  printf "%s\n" "  into files <blockname>.bas"
  printf "%s\n" ""
  exit 0
 else
  printf "%s" "ExtractAllBas.sh -h|<xbamodule>"
  printf "%s\n" " unrecognized option '$P1' - exiting."
  exit 1
 fi		# have -h
fi	# have -
if [ -z "$P1" ];then
 echo "ExtractAllBas -h|<xbamodule> missing parameter(s) - abandoned."
 read -p "Enter ctrl-c to remain in Terminal: "
 exit 1
fi
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

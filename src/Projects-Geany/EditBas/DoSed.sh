#!/bin/bash
# DoSed.sh - Run *sed to fix Make..Bas.tmp > Make..Bas.
#	4/19/25.	wmk.
#
# Usage. bash DoSed.sh -h|<basmodule> <xbafile>
#
#	-h = only display DoSed shell help
#	<basmodule> = name of .bas module to extract
#	<xbafile> = .xba from which to extract; (e.g. Module1)
#
# Exit. *projpath/sedextbas.tmp > sedextbas.txt
#		*projpath/awkxba2xba1.tmp > awkxba2xba1.txt
#		*projpath/MakeExtractBas.tmp >  $projpath/MakeExtractBas
#		*projpath/MakeDeleteXBAbas.tmp >  $projpath/MakeDeleteXBAbas
#		*projpath/MakeReplaceBas.tmp >  $projpath/MakeReplaceBas
#
# Modification History.
# ----------------------
# 4/17/25.	wmk.	-h option support.
# 4/19/25.	wmk.	(automated) Modification History sorted.
# 4/19/25.	wmk.	updated. 
# 1/13/24.	wmk.	code checked for BadAss/src. 
# 11/12/23.	wmk.	(automated) Version 3.0.6 Make old paths removed. 
# 8/23/23.	wmk.	ver2.0 change to use /src as parent folder. 
# 8/13/23.	wmk.	edited for MNcrwg44586; change to *sed from *awk for 
# 8/13/23.	 extraction from .xba file. 
# 6/25/23.	wmk.	*projpath corrected. 
# 6/24/23.	wmk.	modified for FLsara86777; added awkxba2xba1.tmp to edits. 
# 3/8/22.	wmk.	original code. 
#
# P1=-h|<basmodule>, P2=<xbafile>
#
P1=$1
P2=$2
# -h option code
if [ "${P1:0:1}" == "-" ];then
 option=${P1,,}
 if [ "$option" == "-h" ];then
  printf "%s\n" "DoSed - Run *sed to fix Make..Bas.tmp > Make..Bas."
  printf "%s\n" "DoSed.sh -h|<basmodule> <xbafile>"
  printf "%s\n" ""
  printf "%s\n" "  -h = only display DoSed shell help"
  printf "%s\n" "  <basmodule> = name of .bas module to extract"
  printf "%s\n" "  <xbafile> = .xba from which to extract; (e.g. Module1)"
  printf "%s\n" ""
  printf "%s\n" "Results: *projpath/sedextbas.tmp > sedextbas.txt"
  printf "%s\n" "	*projpath/awkxba2xba1.tmp > awkxba2xba1.txt"
  printf "%s\n" "   *projpath/MakeExtractBas.tmp >  $projpath/MakeExtractBas"
  printf "%s\n" "	*projpath/MakeDeleteXBAbas.tmp >  $projpath/MakeDeleteXBAbas"
  printf "%s\n" "	*projpath/MakeReplaceBas.tmp >  $projpath/MakeReplaceBas"
  printf "%s\n" ""
  exit 0
 else
  printf "%s" "DoSed.sh -h|<basmodule> <xbafile>"
  printf "%s\n" " unrecognized option '$P1' - exiting."
  exit 1
 fi		# have -h
fi	# have -
if [ -z "$P1" ] || [ -z "$P2" ];then
 printf "%s" "DoSed.sh -h|<basmodule> <xbafile>"
 printf "%s" " missing parameter(s) - abandoned."
 exit 1
 read -p "Enter ctrl-c to remain in Terminal: "
fi
#
projpath=$libbase/src/Projects-Geany/EditBas
echo $PWD
printf "%s\n" "s?<basmodule>?$P1?g" > $projpath/sedatives.txt 
printf "%s\n" "s?<xbafile>?$P2?g" >> $projpath/sedatives.txt
#echo "s?<basmodule>?$P1?g" > $projpath/sedatives.txt
#echo "s?<xbafile>?$P2?g" >> $projpath/sedatives.txt
sed -f $projpath/sedatives.txt  $projpath/sedextbas.tmp > sedextbas.txt
sed -f $projpath/sedatives.txt  $projpath/awkxba2xba1.tmp > awkxba2xba1.txt
sed -f $projpath/sedatives.txt  $projpath/MakeExtractBas.tmp >  $projpath/MakeExtractBas
sed -f $projpath/sedatives.txt  $projpath/MakeDeleteXBAbas.tmp >  $projpath/MakeDeleteXBAbas
sed -f $projpath/sedatives.txt  $projpath/MakeReplaceBas.tmp >  $projpath/MakeReplaceBas
echo "DoSed (EditBas) $P1 $P2 complete."
# end DoSed.sh

#!/bin/bash
# DoSed.sh - Run *sed to fix Make..Bas.tmp > Make..Bas.
#	1/13/24.	wmk.
#
# Usage. bash DoSed.sh <basmodule> <xbafile>
#
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
# 8/23/23.	wmk.	ver2.0 change to use /src as parent folder.
# 11/12/23.	wmk.	(automated) Version 3.0.6 Make old paths removed.
# 1/13/24.	wmk.	code checked for BadAss/src.
# Legacy mods.
# 8/13/23.	wmk.	edited for MNcrwg44586; change to *sed from *awk for
#			 extraction from .xba file.
# Legacy mods.
# 6/24/23.	wmk.	modified for FLsara86777; added awkxba2xba1.tmp to edits.
# 6/25/23.	wmk.	*projpath corrected.
# Legacy mods.
# 3/8/22.	wmk.	original code.
# P1=<basmodule>, P2=<xbafile>
P1=$1
P2=$2
if [ -z "$P1" ] || [ -z "$P2" ];then
 echo "DoSed <basmodule> <xbafile> missing parameter(s) - abandoned."
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

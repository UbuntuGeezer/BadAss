#!/bin/bash
# DoSedBuild.sh - Run sed to fix BadAss../<xbafile>/MakeBuildLib.tmp > MakeBuildLib.
#	4/19/25.	wmk.
#
# Usage. bash DoSedBuild.sh -h|<xbafile>
#
#	-h = only display DoSedBuild shell help
#	<xbafile> = .xba to build
#
# Entry. *projpath/MakeBuildLib.tmp = MakeBuildLib template
#		contains line with 'insertbaslist'
#
# Exit.	*libbase/src/Basic/<xbafile>/MakeBuildLib = makefile to build
#			<xbafile>.xba
#
# Modification History.
# ----------------------
# 4/19/25.	wmk.	(automated) Modification History sorted.
# 4/19/25.	wmk.	-h option support. 
# 11/12/23.	wmk.	(automated) Version 3.0.9 *libpath introduced. 
# 11/12/23.	wmk.	(automated) Updated for BadAss library (Lenovo). 
# 8/31/23.	wmk.	*folderbase definition made unconditional for HPPavilion 
# 8/31/23.	 branch of project. 
# 8/24/23.	wmk.	ver2.0 mods for using /src as parent folder. 
# 8/22/23.	wmk.	modified for FLsara86777 libary. 
# 8/21/23.	wmk.	*codebase, *pathbase corrected and unconditional; 
# 8/21/23.	 MNcrwg44586 added to comments. 
# 6/27/23.	wmk.	*mawk changed to *gawk; awkbaslist1 added. 
# 6/25/23.	wmk.	original code; edited for FLsara86777 library; mod to 
# 6/25/23.	 rebuild <xbafile>Bas.txt to account for new .bas files. 
# 6/20/23.	wmk.	original code; adapted from DoSed for Territories. 
# 9/23/22.	wmk.	(automated) CB *codebase env var support. 
# 4/24/22.	wmk.	*pathbase* env var included. 
# 3/8/22.	wmk.	original code. 
#
# P1=<xbafile>
#
P1=$1
# -h option code
if [ "${P1:0:1}" == "-" ];then
 option=${P1,,}
 if [ "$option" == "-h" ];then
  printf "%s\n" "DoSedBuild - Run sed to fix BadAss../<xbafile>/MakeBuildLib.tmp > MakeBuildLib."
  printf "%s\n" "DoSedBuild.sh -h|<xbafile>"
  printf "%s\n" ""
  printf "%s\n" "  -h = only display DoSedBuild shell help"
  printf "%s\n" "  <xbafile> = .xba to build (e.g.Module1)"
  printf "%s\n" ""
  printf "%s\n" "Results: *libbase/src/Basic/<xbafile>/MakeBuildLib = makefile to build"
  printf "%s\n" "   <xbafile>.xba"
  printf "%s\n" ""
  exit 0
 else
  printf "%s" "DoSedBuild.sh -h|<xbafile>"
  printf "%s\n" " unrecognized option '$P1' - exiting."
  exit 1
 fi		# have -h
fi	# have -
if [ -z "$P1" ];then
 printf "%s" "DoSedBuild.sh -h|<xbafile>"
 printf "%s\n" "EditBas/DoSedBuild -h|<xbafile> missing parameter(s) - abandoned."
 read -p "Enter ctrl-c to remain in Terminal: "
 exit 1
fi
echo "** WARNING: Ensure that any NEW .bas files have been added into"
echo " <xbafile>Bas.txt (MakeAddBas) before continuing..."
read -p "  OK to continue (y/n)? "
yn=${REPLY^^}
if [ "$yn" != "Y" ];then
 echo "DoSedBuild abandoned at user request."
 read -p "Enter ctrl-c to remain in Terminal: "
 exit 0
fi
#
projpath=$libbase/src/Projects-Geany/EditBas
targpath=$libbase/src/Basic/$P1
fsuffx=Bas
#mawk -f $projpath/awkbaslist.txt $targpath/$P1$fsuffx.txt > baslistvar.txt
gawk -f $projpath/awkbaslist.txt $targpath/$P1$fsuffx.txt > $TEMP_PATH/baslistvar.txt
printf '%s\n' "export last_line=\\" > scratch.sh
wc -l $TEMP_PATH/baslistvar.txt | mawk '{print $1}' >> scratch.sh
#cat scratch.sh
chmod +x scratch.sh
. ./scratch.sh
#read -p "Enter ctrl-c to exit DoSedBuild: "
gawk -v endline=$last_line -f $projpath/awkbaslist1.txt $TEMP_PATH/baslistvar.txt > baslistvar.txt
#read -p "Enter ctrl-c to exit DoSedBuild: "
sed '/insertbaslist/r baslistvar.txt' $projpath/MakeBuildLib.tmp > $targpath/MakeBuildLib
sed -i "s?<xbafile>?$P1?g" $targpath/MakeBuildLib
#mawk -f $projpath/awkbaslist.txt $projpath/MakeBuildLib.tmp \
#  > $targpath/MakeBuildLib 
#echo "s?<xbafile>?$P1?g" > $projpath/sedatives.txt
#sed -i -f $projpath/sedatives.txt  $targpath/MakeBuildLib
echo "DoSedBuild $P1 complete."
# end DoSedBuild.sh

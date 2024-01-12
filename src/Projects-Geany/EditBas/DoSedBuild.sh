#!/bin/bash
# 2023-11-12.	wmk.	(automated) Version 3.0.6 paths eliminated (HPPavilion).
# 2023-11-12.	wmk.	(automated) Version 3.0.9 *libpath introduced (HPPavilion).
# DoSedBuild.sh - Run sed to fix FLsara86777 MakeBuildLib.tmp > MakeBuildLib.
#	11/12/23.	wmk.
#
# Usage. bash DoSedBuild.sh <xbafile>
#
#	<xbafile> = .xba to build
#
# Entry. *projpath/MakeBuildLib.tmp = MakeBuildLib template
#		contains line with 'insertbaslist'
#
# Exit.	*pathbase/Basic/<xbafile>/MakeBuildLib = makefile to build
#			<xbafile>.xba
#
# Modification History.
# ----------------------
# 8/24/23.	wmk.	ver2.0 mods for using /src as parent folder.
# 8/31/23.	wmk.	*folderbase definition made unconditional for HPPavilion
#			 branch of project.
# 11/12/23.	wmk.	(automated) Version 3.0.6 Make old paths removed.
# 11/12/23.	wmk.	(automated) Version 3.0.9 *libpath introduced.
# Legacy mods.
# 8/22/23.	wmk.	modified for FLsara86777 libary.
# Legacy mods.
# 6/25/23.	wmk.	original code; edited for FLsara86777 library; mod to
#			 rebuild <xbafile>Bas.txt to account for new .bas files.
# 6/27/23.	wmk.	*mawk changed to *gawk; awkbaslist1 added.
# 8/21/23.	wmk.	*codebase, *pathbase corrected and unconditional;
#			 MNcrwg44586 added to comments.
# Legacy mods.
# 6/20/23.	wmk.	original code; adapted from DoSed for Territories.
# Legacy mods.
# 3/8/22.	wmk.	original code.
# 4/24/22.	wmk.	*pathbase* env var included.
# 9/23/22.  wmk.    (automated) CB *codebase env var support.
export codebase=$folderbase/GitHub/Libraries-Project/FLsara86777/src
export pathbase=$codebase
#
P1=$1
if [ -z "$P1" ];then
 echo "EditBas/DoSedBuild <xbafile> missing parameter(s) - abandoned."
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

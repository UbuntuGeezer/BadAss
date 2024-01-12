#!/bin/bash
# 2023-11-12.	wmk.	(automated) Version 3.0.6 paths eliminated (HPPavilion).
# 2023-11-12.	wmk.	(automated) Version 3.0.9 *libpath introduced (HPPavilion).
# DoSedAdd.sh - Run sed to fix MakeAddLib.tmp > MakeAddLib.
#	11/12/23.	wmk.
#
# Usage. bash DoSedAdd.sh <xbafile> <basname> [<beforebas>]
#
#	<xbafile> = .xba to build
#	<basname> = name of .bas to add (e.g. NewSub)
#	<beforebas> = (optional) name of .bas at which to insert <basname> before
#
# Entry. *pathbase/Basic/<xbafile/<xbafile>Bas.txt = list of .bas blocks in <xbafile.xba>
#
# Exit.	*pathbase/Basic/<xbafile>/<xbafile>Bas.txt = new list of .bas blocks
#			for building <xbafile>.xba
#
# Modification History.
# ----------------------
# 11/12/23.	wmk.	(automated) Version 3.0.6 Make old paths removed.
# 11/12/23.	wmk.	(automated) Version 3.0.9 *libpath introduced.
# Legacy mods.
# 6/25/23.	wmk.	original code; edited for FLsara86777 library; mod to
#			 rebuild <xbafile>Bas.txt to account for new .bas files.
# Legacy mods.
# 6/20/23.	wmk.	original code; adapted from DoSed for Territories.
# Legacy mods.
# 3/8/22.	wmk.	original code.
# 4/24/22.	wmk.	*pathbase* env var included.
# 9/23/22.  wmk.    (automated) CB *codebase env var support.
#
P1=$1
P2=$2
P3=$3
if [ -z "$P1" ] || [ -z "$P2" ];then
 echo "EditBas/DoSedAdd <xbafile> <basname> [<beforebas>] missing parameter(s) - abandoned."
 read -p "Enter ctrl-c to remain in Terminal: "
 exit 1
fi
#
projpath=$libbase/src/Projects-Geany/EditBas
targpath=$libbase/src/Basic/$P1
fsuffx=Bas
sed -n "/^$P2\$/p" $targpath/$P1$fsuffx.txt > $TEMP_PATH/CurrBasList.txt
if test -s $TEMP_PATH/CurrBasList.txt;then
 echo "** block with same name '$P2' already in $P1$suffx - skipping.. "
  read -p "Enter ctrl-c to remain in Terminal: "
  exit 0
fi
grep -e "$P2" $targpath/$P1$fsuffx.txt > $TEMP_PATH/CurrBasList.txt
if test -s $TEMP_PATH/CurrBasList.txt;then
 echo "** blocks with similar name(s) to '$P2' already in $P1$suffx.. "
 cat $TEMP_PATH/CurrBasList.txt
 read -p "Do you wish to continue (y/n)? "
 yn=${REPLY^^}
 if [ "$yn" != "Y" ];then
  echo "DoSedAdd.sh abandoned at user request."
  read -p "Enter ctrl-c to remain in Terminal: "
  exit 0
 fi
fi
#echo "early termination for testing..."
#read -p "Enter ctrl-c to remain in Terminal: "
#exit 0
#
sed  "s?<xbafile>?$P1?g;s?<basname>?$P2?g;s?<where>?$P3?g" $projpath/MakeAddBas.tmp > $projpath/MakeAddBas
#mawk -f $projpath/awkbaslist.txt $projpath/MakeAddLib.tmp \
#  > $targpath/MakeAddLib 
#echo "s?<xbafile>?$P1?g" > $projpath/sedatives.txt
#sed -i -f $projpath/sedatives.txt  $targpath/MakeAddLib
echo "DoSedAdd complete."
# end DoSedAdd.sh

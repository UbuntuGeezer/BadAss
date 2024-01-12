#!/bin/bash
# 2023-11-12.	wmk.	(automated) Version 3.0.6 paths eliminated (HPPavilion).
# 2023-11-12.	wmk.	(automated) Version 3.0.9 *libpath introduced (HPPavilion).
# DoSedDel.sh - Run sed to fix MakeDelBas.tmp > MakeDelBas.
#	11/12/23.	wmk.
#
# Usage. bash DoSedDel.sh <xbafile> <basname>]
#
#	<xbafile> = .xba to build
#	<basname> = name of .bas to add (e.g. NewSub)
#
# Entry. *pathbase/Basic/<xbafile/<xbafile>Bas.txt = list of .bas blocks in <xbafile.xba>
#
# Exit.	*pathbase/Basic/<xbafile>/<xbafile>Bas.txt = new list of .bas blocks
#			for building <xbafile>.xba
#
# Modification History.
# ----------------------
# 6/25/23.	wmk.	orignal code; adpapted from DoSedAdd.
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
 echo "EditBas/DoSedDel <xbafile> <basname> missing parameter(s) - abandoned."
 read -p "Enter ctrl-c to remain in Terminal: "
 exit 1
fi
#
projpath=$libbase/src/Projects-Geany/EditBas
targpath=$libbase/src/Basic/$P1
sed  "s?<xbafile>?$P1?g;s?<basname>?$P2?g" $projpath/MakeDelBas.tmp > $projpath/MakeDelBas
echo "DoSedDel $P1 $P2 complete."
# end DoSedDel.sh

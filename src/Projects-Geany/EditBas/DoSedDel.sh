#!/bin/bash
# DoSedDel.sh - Run sed to fix MakeDelBas.tmp > MakeDelBas.
#	4/19/25.	wmk.
#
# Usage. bash DoSedDel.sh -h|<xbafile> <basname>]
#
#	-h = only display DoSedDel shell help
#   <xbafile> = .xba to remove .bas block from (e.g. Module1)"
#	<basname> = name of .bas to add (e.g. NewSub)"
#
# Entry. *pathbase/Basic/<xbafile/<xbafile>Bas.txt = list of .bas blocks in <xbafile.xba>
#
# Exit.	*pathbase/Basic/<xbafile>/<xbafile>Bas.txt = new list of .bas blocks
#			for building <xbafile>.xba
#
# Modification History.
# ----------------------
# 4/19/25.	wmk.	(automated) Modification History sorted.
# 4/17/25.	wmk.	-h option support. 
# 11/12/23.	wmk.	(automated) Version 3.0.9 *libpath introduced. 
# 11/12/23.	wmk.	(automated) Version 3.0.6 Make old paths removed. 
# 6/25/23.	wmk.	orignal code; adpapted from DoSedAdd. 
# 6/25/23.	wmk.	original code; edited for FLsara86777 library; mod to 
# 6/25/23.	 rebuild <xbafile>Bas.txt to account for new .bas files. 
# 6/20/23.	wmk.	original code; adapted from DoSed for Territories. 
# 9/23/22.	wmk.	(automated) CB *codebase env var support. 
# 4/24/22.	wmk.	*pathbase* env var included. 
# 3/8/22.	wmk.	original code. 
#
P1=$1
P2=$2
P3=$3
# -h option code
if [ "${P1:0:1}" == "-" ];then
 option=${P1,,}
 if [ "$option" == "-h" ];then
  printf "%s\n" "DoSedDel - Run sed to fix MakeDelBas.tmp > MakeDelBas."
  printf "%s\n" "DoSedDel.sh -h|<xbafile> <basname>]"
  printf "%s\n" ""
  printf "%s\n" "  -h = only display DoSedDel shell help"
  printf "%s\n" "  <xbafile> = .xba to remove .bas block from (e.g. Module1)"
  printf "%s\n" "  <basname> = name of .bas to add (e.g. NewSub)"
  printf "%s\n" ""
  printf "%s\n" "Results: *pathbase/Basic/<xbafile>/<xbafile>Bas.txt = new list of .bas blocks"
  printf "%s\n" "  for building <xbafile>.xba"
  printf "%s\n" ""
  exit 0
 else
  printf "%s" "DoSedDel.sh -h|<xbafile> <basname>]"
  printf "%s\n" " unrecognized option '$P1' - exiting."
  exit 1
 fi		# have -h
fi	# have -
if [ -z "$P1" ] || [ -z "$P2" ];then
 echo "EditBas/DoSedDel -h|<xbafile> <basname> missing parameter(s) - abandoned."
 read -p "Enter ctrl-c to remain in Terminal: "
 exit 1
fi
#
projpath=$libbase/src/Projects-Geany/EditBas
targpath=$libbase/src/Basic/$P1
sed  "s?<xbafile>?$P1?g;s?<basname>?$P2?g" $projpath/MakeDelBas.tmp > $projpath/MakeDelBas
echo "DoSedDel $P1 $P2 complete."
# end DoSedDel.sh

#!/bin/bash
# ExtractAllBas.sh - Extract all .bas modules from .xba.
#	8/25/22.	wmk.
#
# Usage. bash  ExtractAllBas.sh <baslist> <xbafile>
#
#	<baslist> = filename containing list of all .bas modules within .xba
#				 (e.g. Module2List.txt)
#	<xbafile> = .xba filename (e.g. Module2)
#
# Exit. all .bas modules within .xba file extracted to separate files
#		 modules delimited by ' ..<module-name>.bas and '/**/ lines.
#
P1=$1
P2=$2
if [ -z "$P1" ] || [ -z "$P2" ];then
 echo "ExtractAllBas <baslist> <xbafile> missing parameter(s) - abandoned."
 exit 1
fi
projpath=$pathbase/Projects-Geany/EditBas
cd $projpath
# loop on <baslist> lines
file=$projpath/$P1
while read -e; do
  len=${#REPLY}
  len1=$((len-1))
  mod_name=${REPLY:0:len}
  echo "Processing module $mod_name..."
  ./DoSed.sh $mod_name $P2.xba
  make -f MakeExtractBas
done < $file
echo "  ExtractAllBas complete."
#end ExtractAllBas.sh


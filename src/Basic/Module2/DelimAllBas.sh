#!/bin/bash
echo " ** DelimAllBas.sh out-of-date **";exit 1
# DelimAllBas.sh - Delim all .bas modules from .xba.
#	8/25/22.	wmk.
#
# Usage. bash  DelimAllBas.sh <baslist> <xbafile>
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
 echo "DelimAllBas <baslist> <xbafile> missing parameter(s) - abandoned."
 exit 1
fi
projpath=$pathbase/Projects-Geany/BadAssLibrary/Module2
cd $projpath
#sed -i -f sedAddBasEnd.txt Module2.xba
#
if [ 1 -eq 1 ];then
# loop on <baslist> lines
file=$projpath/$P1
while read -e; do
  len=${#REPLY}
  len1=$((len-1))
  mod_name=${REPLY:0:len}
  echo "Processing module $mod_name..."
  sed -i -f sedAddBasEnd.txt $mod_name
  #make -f MakeDelimBas
done < $file
fi
#
echo "  DelimAllBas complete."
#end DelimAllBas.sh


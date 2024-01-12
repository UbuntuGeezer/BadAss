#!/bin/bash
# 2023-11-13.	wmk.	(automated) Version 3.0.6 paths eliminated (HPPavilion).
# 2023-11-13.	wmk.	(automated) Version 3.0.9 *libpath introduced (HPPavilion).
# ReplaceXBA.sh - replace .bas into .xba file.
#	3/8/22.
# modulename P1
# xbafile P2
P1=$1
P2=$2
cp $1.bas $TEMP_PATH/scratch.bas
sed -i "/\/\ $P1.bas/d;/\/\*\*\//d" $TEMP_PATH/scratch.bas
mawk "/\/\/ $P1.bas/{f=1;print;while(getline < \"$TEMP_PATH/scratch.bas\"){print}}/\/\*\*\//{f=0}!f" $2 > new$P2
sed -i "s?\'?\&apos\;?g;s?\"?\&quot\;?g" new$P2

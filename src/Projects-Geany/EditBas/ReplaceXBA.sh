#!/bin/bash
# ReplaceXBA.sh - replace .bas into .xba file.
#	4/19/25.
#
# Usage. ReplaceXBA.sh -h|<modulename> <xbafile>
#
#	-h = only display ReplaceXBA shell help
#	<modulename> = .bas block name
#	<xbafile> = .xba filename (e.g. Module1)
#
# Exit. src/Basic/<modulename>/new<modulename>.xba = .xba with new .bas block added. 
#
# Modification History.
# ---------------------
# 4/19/25.	wmk.	-h option added.
# 11/13/23.	wmk.	(automated) Version 3.0.9 *libpath introduced (HPPavilion).
# 3/8/22.	wmk.	original.
#
# P1=<modulename>, P2=<xbafile>
#
P1=$1
P2=$2
# -h option code
if [ "${P1:0:1}" == "-" ];then
 option=${P1,,}
 if [ "$option" == "-h" ];then
  printf "%s\n" "ReplaceXBA - replace .bas into .xba file."
  printf "%s\n" "ReplaceXBA.sh -h|<modulename> <xbafile>"
  printf "%s\n" ""
  printf "%s\n" "  -h = only display ReplaceXBA shell help"
  printf "%s\n" "  <modulename> = .bas block name"
  printf "%s\n" "  <xbafile> = .xba filename (e.g. Module1)"
  printf "%s\n" ""
  printf "%s\n" "Results:"
  printf "%s\n" ""
  exit 0
 else
  printf "%s" "ReplaceXBA.sh -h|<modulename> <xbafile>"
  printf "%s\n" " unrecognized option '$P1' - exiting."
  exit 1
 fi		# have -h
fi	# have -
if [ -z "$P1" ] || [ -z "$P2" ];then
 printf "%s\n" "ReplaceXBA -h|<modulename> <xbafile> missing parameter(s) - abandoned."
 exit 1
fi
cp $1.bas $TEMP_PATH/scratch.bas
sed -i "/\/\ $P1.bas/d;/\/\*\*\//d" $TEMP_PATH/scratch.bas
mawk "/\/\/ $P1.bas/{f=1;print;while(getline < \"$TEMP_PATH/scratch.bas\"){print}}/\/\*\*\//{f=0}!f" $2 > new$P2.xba
sed -i "s?\'?\&apos\;?g;s?\"?\&quot\;?g" new$P2.xba

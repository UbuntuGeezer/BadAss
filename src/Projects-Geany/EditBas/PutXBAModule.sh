#!/bin/bash
# PutXBAModule - Copy src/Basic/<xbamodule>/<xbamodule>.xba module to Release folder.
#	4/19/25.	wmk.
#
# Usage. bash  PutXBAModule  <xbamodule>
#
#	-h = only display PutXBAModule shell help
#	<xbamodule> = module name to "put" to /Release (e.g. Module1)
#
# Exit. Edit/<xbamodule>/<xbamodule>.xba copied to ../Release
#
# Modification History.
# ---------------------
# 4/19/25.	wmk.	(automated) Modification History sorted.
# 4/19/25.	wmk.	updated. 
# 4/19/25.	wmk.	-h option support. 
# 11/13/23.	wmk.	(automated) Version 3.0.9 *libpath introduced. 
# 11/13/23.	wmk.	(automated) Version 3.0.6 Make old paths removed. 
# 6/25/23.	wmk.	NEW .bas block warning added. 
# 6/24/23.	wmk.	modified for FLsara86777. 
# 3/8/22.	wmk.	original code; adapted from GetXBAModule. 
#
# P1=<xbamodule>
#
projbase=$libbase/src/Projects-Geany/EditBas
gitbase=$libbase
P1=$1
# -h option code
if [ "${P1:0:1}" == "-" ];then
 option=${P1,,}
 if [ "$option" == "-h" ];then
  printf "%s\n" "PutXBAModule - Copy /Basic/<xbamodule>/<xbamodule>.xba module to Release folder."
  printf "%s\n" "PutXBAModule  -h|<xbamodule>"
  printf "%s\n" ""
  printf "%s\n" "  -h = only display PutXBAModule shell help"
  printf "%s\n" "  <xbamodule> = module name to "put" to /Release (e.g. Module1)"
  printf "%s\n" ""
  printf "%s\n" "Results: Basic/<xbamodule>/<xbamodule>.xba copied to ../Release"
  printf "%s\n" ""
  exit 0
 else
  printf "%s" "PutXBAModule  -h|<xbamodule>"
  printf "%s\n" " unrecognized option '$P1' - exiting."
  exit 1
 fi		# have -h
fi	# have -
if [ -z "$P1" ];then
 echo "PutXBAModule -h|<xbamodule> missing parameter(s) - PutXBAModule abandoned.**"
 exit 1
fi
cd $gitbase
if ! test -f $libbase/src/Basic/$P1/$P1.xba;then
 echo "** file Basic/$P1.xba not found for copy - PutXBAModule abandoned.**"
 exit 1
fi
echo "** WARNING: Ensure that any NEW .bas files have been added into"
echo " <xbafile>Bas.txt (MakeAddBas) before continuing..."
read -p "  OK to continue (y/n)? "
yn=${REPLY^^}
if [ "$yn" != "Y" ];then
 echo "PutXBAModule abandoned at user request."
 read -p "Enter ctrl-c to remain in Terminal: "
 exit 0
fi
if [ $libbase/src/Release/$P1.xba -nt $gitbase/Basic/$P1/$P1.xba ];then
 echo "**$P1.xba copy skipped - Basic/$P1.xba is not newer.**"
 ~/sysprocs/LOGMSG "**$ PutXBAModule $P1.xba copy skipped - Basic/$P1.xba is newer.**"
else
 cp -u $libbase/src/Basic/$P1/$P1.xba $P1.xba 
 echo "Basic/$P1.xba copied to GitHub project folder."
 ~/sysprocs/LOGMSG "  PutXBAModule Basic/$P1.xba copied to GitHub project folder."
fi
# end PutXBAModule.sh.

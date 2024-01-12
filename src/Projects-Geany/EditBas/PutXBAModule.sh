#!/bin/bash
# 2023-11-13.	wmk.	(automated) Version 3.0.6 paths eliminated (HPPavilion).
# 2023-11-13.	wmk.	(automated) Version 3.0.9 *libpath introduced (HPPavilion).
# PutXBAModule - Copy /Libraries-Project/Territories .xba module to Relese folder.
#	6/25/23.	wmk.
#
# Usage. bash  PutXBAModule  <xbamodule>
#
#	<xbamodule> = module name to "put" to /Release (e.g. Module1)
#
# Exit. Edit/<xbamodule>/<xbamodule>.xba copied to ../Release
#
# Modification History.
# ---------------------
# 11/13/23.	wmk.	(automated) Version 3.0.6 Make old paths removed.
# 11/13/23.	wmk.	(automated) Version 3.0.9 *libpath introduced.
# Legacy mods.
# 6/24/23.	wmk.	modified for FLsara86777.
# 6/25/23.	wmk.	NEW .bas block warning added.
# Legacy mods.
# 3/8/22.	wmk.	original code; adapted from GetXBAModule.
projbase=$libbase/src/Projects-Geany/EditBas
gitbase=$libbase
P1=$1
if [ -z "$P1" ];then
 echo "PutXBAModule <xbamodule> missing parameter(s) - PutXBAModule abandoned.**"
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

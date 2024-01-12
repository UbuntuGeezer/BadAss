#!/bin/bash
# GetBasList.sh - Get .bas block list from FLsara86777 .xba Module.
#	1/12/24.	wmk.
#
# Usage. bash  GetBasList.sh <xbamodule>
#
#	<xbamodule> = .xba module name (e.g. Module1)
#
# Entry. Basic/<xbamodule>/<xbamodule>.xba = .xba module source code from Calc
#	Territories library.
#
# Exit.	Basic/<xbamodule>/<xbamodule>Bas.txt = list of .bas block names within
#	<xbamodule>.xba
#
# Modification History.
# ---------------------
# 1/12/24.	wmk.	code checked for BadAss library compatibility.
# Legacy mods.
# 11/13/23.	wmk.	(automated) Version 3.0.6 Make old paths removed.
# 11/13/23.	wmk.	(automated) Version 3.0.9 *libpath introduced.
# Legacy mods.
# 6/22/23.	wmk.	original shell.
# 8/21/23.	wmk.	*pathbase, *codebase unconditional.
# 8/22/23.	wmk.	*codebase used throughout; FLsara86777 added to comments.
#
# P1=<xbamodule>
#
P1=$1
if [ -z "$P1" ];then
 echo "GetBasList <xbamodule> missing parameter(s) - abandoned."
 read -p "Enter ctrl-c to remain in Terminal: "
 exit 1
fi
#	Environment vars:
#if [ -z "$TODAY" ];then
# . ~/GitHub/TerritoriesCB/Procs-Dev/SetToday.sh
#TODAY=2022-04-22
#fi
#procbodyhere
srcbase=$libbase/src/Basic/$P1
targbase=$libbase/src/Basic/$P1
targf1=Bas.txt
targfile=$P1$targf1
grep -e "//.*\.bas" $srcbase/$P1.xba | mawk -F "/" '{print $3}' > $TEMP_PATH/BasList.txt
mawk -F "." '{if(substr($1,1,1) == " ")print substr($1,2,99);else print $1}' \
 $TEMP_PATH/BasList.txt > $targbase/$targfile
#endprocbody
echo "  GetBasList complete."
echo "  List of .bas routines within $P1.xba on $targfile."
~/sysprocs/LOGMSG "  GetBasList complete."
# end GetBasList.sh

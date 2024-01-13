#!/bin/bash
# ExtractXBAhdr.sh - Extract .xba header block from .xba module.
#	1/13/24.	wmk.
#
# Usage. bash  ExtractXBAhdr.sh <xbamodule>
#
#	<xbamodule> = .xba module to exract .xba header block from
#
# Entry. 
#
# Dependencies.
#
# Modification History.
# ---------------------
# 11/13/23.	wmk.	(automated) UnKillShell to reinstate shell.
# 11/13/23.	wmk.	(automated) Version 3.0.9 *libpath introduced.
# 1/13/24.	wmk.	code checked for BadAss/src.
#
# P1=<xbamodule>
P1=$1
if [ -z "$P1" ];then
 echo "ExtractXBAhdr <xbamodule> missing parameter(s) - abandoned."
 exit 1
fi
#	Environment vars:
#if [ -z "$TODAY" ];then
# . ~/GitHub/TerritoriesCB/Procs-Dev/SetToday.sh
#TODAY=2022-04-22
#fi
#procbodyhere
projpath=$libbase/src/Projects-Geany/EditBas
cd $projpath
xbapath=$libbase/src/Basic/$P1
hdr=Hdr
sed -n '/xml version/,/\*\*\*\*\*  BASIC/p' $xbapath/$P1.xba > $xbapath/$P1$hdr.xba
#endprocbody
echo "  ExtractXBAhdr $P1 $P2 complete."
~/sysprocs/LOGMSG "  ExtractXBAhdr $P1 $P2 complete."
# end ExtractXBAhdr.sh

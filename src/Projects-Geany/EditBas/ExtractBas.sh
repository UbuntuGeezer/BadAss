#!/bin/bash
# ExtractBas.sh - Extract BadAss .bas block from .xba module.
#	1/13/24.	wmk.
#
# Usage. bash  ExtractBas.sh <xbamodule> <basblock>
#
#	<xbamodule> = .xba module to exract <basblock> from
#	<basblock> = .bas block within <xbamodule> to extract
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
# P1=<xbamodule> P2=<basblock>
P1=$1
P2=$2
if [ -z "$P1" ] || [ -z "$P2" ];then
 echo "ExtractBas <xbamodule> <basblock> missing parameter(s) - abandoned."
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
$projpath/DoSed.sh   $P2 $P1
make --silent -f $projpath/MakeExtractBas
#make -f $projpath/MakeExtractBas
#endprocbody
echo "  ExtractBas $P1 $P2 complete."
~/sysprocs/LOGMSG "  ExtractBas $P1 $P2 complete."
# end ExtractBas.sh

#!/bin/bash
# 2023-11-13.	wmk.	(automated) Version 3.0.6 paths eliminated (HPPavilion).
# 2023-11-13.	wmk.	(automated) Version 3.0.9 *libpath introduced (HPPavilion).
# ExtractBas.sh - Extract FLsara86777 .bas block from .xba module.
#	11/13/23.	wmk.
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
# 11/13/23.	wmk.	(automated) Version 3.0.6 Make old paths removed.
# 11/13/23.	wmk.	(automated) Version 3.0.9 *libpath introduced.
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

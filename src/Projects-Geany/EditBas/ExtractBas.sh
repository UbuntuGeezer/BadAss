#!/bin/bash
# ExtractBas.sh - Extract BadAss .bas block from .xba module.
#	7/22/26.	wmk.
#
# Usage. bash  ExtractBas.sh -h|<xbamodule> <basblock>
#
#	-h = only display ExtractBas shell help
#	<xbamodule> = .xba module to exract <basblock> from
#	<basblock> = .bas block within <xbamodule> to extract
#
# Exit. <basblock>.bas extracted from <xbamodule>.xba to Basic/<xbamodule>/
#
# Modification History.
# ---------------------
# 7/22/26.	wmk.	updated for Windows 11.
# 7/22/26.	wmk.	(automated) UnKillShell to reinstate shell.
# 4/19/25.	wmk.	-h option support.
# 11/13/23.	wmk.	(automated) UnKillShell to reinstate shell.
# 11/13/23.	wmk.	(automated) Version 3.0.9 *libpath introduced.
# 1/13/24.	wmk.	code checked for BadAss/src.
#
# P1=-h|<xbamodule> P2=<basblock>
#
P1=$1
P2=$2
# -h option code
if [ "${P1:0:1}" == "-" ];then
 option=${P1,,}
 if [ "$option" == "-h" ];then
  printf "%s\n" "ExtractBas - Extract BadAss .bas block from .xba module."
  printf "%s\n" "ExtractBas.sh -h|<xbamodule> <basblock>"
  printf "%s\n" ""
  printf "%s\n" "  -h = only display ExtractBas shell help"
  printf "%s\n" "  <xbamodule> = .xba module to exract <basblock> from"
  printf "%s\n" "  <basblock> = .bas block within <xbamodule> to extract"
  printf "%s\n" ""
  printf "%s\n" "Results: <basblock>.bas extracted from <xbamodule>.xba to Basic/<xbamodule>/"
  printf "%s\n" ""
  exit 0
 else
  printf "%s" "ExtractBas.sh -h|<xbamodule> <basblock>"
  printf "%s\n" " unrecognized option '$P1' - exiting."
  exit 1
 fi		# have -h
fi	# have -
if [ -z "$P1" ] || [ -z "$P2" ];then
 printf "%s\n" "ExtractBas -h|<xbamodule> <basblock> missing parameter(s) - abandoned."
 exit 1
fi
#procbodyhere
projpath=C:/Users/vncwm/linux/BadAss/src/Projects-Geany/EditBas
cd $projpath
$projpath/DoSed.sh   $P2 $P1
make --silent -f $projpath/MakeExtractBas
#make -f $projpath/MakeExtractBas
#endprocbody
printf "%s\n" "  ExtractBas $P1 $P2 complete."
$sp/LOGMSG "  ExtractBas (BadAss) $P1 $P2 complete."
# end ExtractBas.sh

#!/bin/bash
# ExtractXBAhdr.sh - Extract .xba header block from .xba module.
#	4/19/25.	wmk.
#
# Usage. bash  ExtractXBAhdr.sh -h|<xbamodule>
#
#	-h = only display ExtractXBAhdr shell help
#	<xbamodule> = .xba module to exract .xba header block from
#
# Exit. src/Basic/<xbamodule>/<xbamodule>Hdr.xba = XBA module header
#
# Modification History.
# ---------------------
# 4/19/25.	wmk.	-h option support.
# 11/13/23.	wmk.	(automated) UnKillShell to reinstate shell.
# 11/13/23.	wmk.	(automated) Version 3.0.9 *libpath introduced.
# 1/13/24.	wmk.	code checked for BadAss/src.
#
# P1=-h|<xbamodule>
#
P1=$1
# -h option code
if [ "${P1:0:1}" == "-" ];then
 option=${P1,,}
 if [ "$option" == "-h" ];then
  printf "%s\n" "ExtractXBAhdr - Extract .xba header block from .xba module."
  printf "%s\n" "ExtractXBAhdr.sh -h|<xbamodule>"
  printf "%s\n" ""
  printf "%s\n" "  -h = only display ExtractXBAhdr shell help"
  printf "%s\n" "  <xbamodule> = .xba module to exract .xba header block from"
  printf "%s\n" ""
  printf "%s\n" "Results: src/Basic/<xbamodule>/<xbamodule>Hdr.xba = XBA module header."
  printf "%s\n" ""
  exit 0
 else
  printf "%s" "ExtractXBAhdr.sh -h|<xbamodule>"
  printf "%s\n" " unrecognized option '$P1' - exiting."
  exit 1
 fi		# have -h
fi	# have -
if [ -z "$P1" ];then
 echo "ExtractXBAhdr -h|<xbamodule> missing parameter(s) - abandoned."
 exit 1
fi
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

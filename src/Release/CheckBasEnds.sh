#!/bin/bash
# CheckBasEnds.sh - Check that all Module1.xba macros have terminating '/**/.
#	4/20/25.	wmk.
#
# Usage. bash  CheckBasEnds.sh [-h|<xbamodule>]
#
#	-h = (optional) only display CheckBasEnds shell help
#	<xbamodule> = (optional) module name to check
#	   default = Module1
#
# Entry. Release/<xbamodule>.xba = .bas source to check
# 
# Exit. Release/<xbamodule>.xba blocks checked for '/**/
#
# Modification History.
# ---------------------
# 4/20/25.	wmk.	add *projpath for base folder allowing run from other folders.
# 4/18/25.	wmk.	-h option support.
# 4/18/25.	wmk.	(automated) build level 4.0.14.
# 8/24/23.	wmk.	original shell.
#
# [P1=-h|<xbamodule>
#
P1=$1
# -h option code
if [ "${P1:0:1}" == "-" ];then
 option=${P1,,}
 if [ "$option" == "-h" ];then
  printf "%s\n" "CheckBasEnds - Check that all Module1.xba macros have terminating '/**/."
  printf "%s\n" "CheckBasEnds [-h|<xbamodule>]"
  printf "%s\n" ""
  printf "%s\n" "  -h = only display CheckBasEnds shell help"
  printf "%s\n" "  <xbamodule> = (optional) module name to check"
  printf "%s\n" "     default = Module1"
  printf "%s\n" ""
  printf "%s\n" "Results: Release/<xbamodule>.xba blocks checked for '/**/"
  printf "%s\n" ""
  exit 0
 else
  printf "%s" "CheckBasEnds - Check that all Module1.xba macros have terminating '/**/."
  printf "%s\n" " unrecognized option '$P1' - exiting."
  exit 1
 fi		# have -h
fi	# have -
if [ -z "$P1" ];then
 P1=Module1
fi
#procbodyhere
projpath=$libbase/src/Release
mawk -f $projpath/awkListBasEnds.txt $P1.xba
#endprocbody
printf "%s\n" "  CheckBasEnds complete."
~/sysprocs/LOGMSG "  CheckBasEnds complete."
# end CheckBasEnds.sh

#!/bin/bash
# CreateBasList.sh - Create <module-name>Bas.txt list in Basic/<module-name> folder.
#	4/15/25.	wmk.
#
# Usage. bash  CreateBasList.sh -h|<module-name>
#
#	-h = only display CreateBasList shell help
#	<module-name> = module name (e.g. Module1)
#
# Entry. ../src/Basic/<module-name> has .bas source files
#	*libbase = FLsara86777lib base folder path
#
# Exit.	../src/Basic/<module-name>/<module-name>Bas.txt = list of .bas files
#	without .bas suffixes.
#
# Modification History.
# ---------------------
# 4/15/25.	wmk.	-h option added.
# 4/15/25.	wmk.	(automated) build level 4.0.14.
# 6/5/24.	wmk.	(automated) printf "%s\n",s to printf,s throughout.
# 6/5/24.	wmk.	(automated) build level 4.0.8.
# 6/5/24.	wmk.	(automated) mods for build level 4.0.8.
# 4/25/24.	wmk.	original shell.
#
# Notes. <module-name>Bas.txt is used by ExtractAllBas to extract all .bas
# macros from the .xba file. It is also used by BuildLib to construct the
# library from source code files.
#
# P1=-h|<module-name>
#
P1=$1
# -h option code
if [ "${P1:0:1}" == "-" ];then
 option=${P1,,}
 if [ "$option" == "-h" ];then
  printf "%s\n" "CreateBasList - Create <module-name>Bas.txt list in Basic/<module-name> folder."
  printf "%s\n" "CreateBasList.sh -h|<module-name>"
  printf "%s\n" ""
  printf "%s\n" "  -h = only display CreateBasList shell help"
  printf "%s\n" "  <module-name> = module name (e.g. Module1)"
  printf "%s\n" ""
  printf "%s\n" "Results:"
  printf "%s\n" "../src/Basic/<module-name>/<module-name>Bas.txt = list of .bas files"
  printf "%s\n" "without .bas suffixes."
  printf "%s\n" ""
  exit 0
 else
  printf "%s" "CreateBasList.sh -h|<module-name>"
  printf "%s\n" " unrecognized option '$P1' - exiting."
  exit 1
 fi		# have -h
fi	# have -
if [ -z "$P1" ];then
 printf "%s\n" "CreateBasList -h|<module-name> missing parameter(s) - abandoned."
 exit 1
fi
~/sysprocs/LOGMSG "  CreateBasList - initiated from Terminal"
printf "%s\n" "  CreateBasList - initiated from Terminal"
# procbodyhere
cd $libbase/src/Basic/$P1
ls \
 *.bas | mawk -F "/" '{print substr($NF,1,length($NF)-4)}' > ${P1}Bas.txt
#endprocbody
printf "%s\n" " ** Reminder: edit ${P1}Bas.txt moving ${P1}Hdr, publicsMM, ${P1}Common, "
printf "%s\n" "    and Main to the top of the list, after removing Module1 from the list."
printf "%s\n" "  CreateBasList complete."
~/sysprocs/LOGMSG "  CreateBasList complete."
# end CreateBasList.sh

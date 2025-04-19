#!/bin/bash
# CopyXBAoverBAS - Copy BadAss Import module .xba code over to Basic folder.
#	4/19/25.	wmk.
#
# Usage.	bash CopyXBAoverBAS.sh -h|<basmodule>
#
#	-h = only display CopyXBAoverBAS shell help
#	<basmodule> = filename of project source .xba to copy (e.g. Module1)
#
# Exit.	/Basic/<basmodule/<basmodule>.xba updated from Import/<basmodule>.xba
#
# Modification History.
# ---------------------
# 4/19/25.	wmk.	(automated) Modification History sorted.
# 4/19/25.	wmk.	-h option support. 
# 1/12/24.	wmk.	code checked for Accounting/BadAss compatibility. 
# 11/13/23.	wmk.	*srcpath corrected. 
# 11/12/23.	wmk.	(automated) Version 3.0.9 *libpath introduced. 
# 11/12/23.	wmk.	(automated) Version 3.0.6 Make old paths removed. 
# 9/17/23.	wmk.	Exit path corrected. 
# 8/23/23.	wmk.	ver2.0 change to use /src as parent folder. 
# 8/22/23.	wmk.	edited for FLsara86777; *codebase used throughout. 
# 8/13/23.	wmk.	*pathbase, *codebase unconditional; corrected to point to 
# 8/13/23.	 GitHub/Libraries-Project/MNcrwg44586.
# 6/25/23.	wmk.	original code; adpated from Territories. 
# 9/23/22.	wmk.	(automated) CB *codebase env var support. 
# 4/24/22.	wmk.	*pathbase* env var included. 
# 3/8/22.	wmk.	description clarified. 
# 3/7/22.	wmk.	original code. 
#
# P1=-h|<xbamodule>
#
P1=$1
srcpath=$libbase/src/Import
targpath=$libbase/src/Basic/$P1
# -h option code
if [ "${P1:0:1}" == "-" ];then
 option=${P1,,}
 if [ "$option" == "-h" ];then
  printf "%s\n" "CopyXBAoverBAS - Copy BadAss Import module .xba code over to Basic folder."
  printf "%s\n" "CopyXBAoverBAS.sh -h|<basmodule>"
  printf "%s\n" ""
  printf "%s\n" "  -h = only display CopyXBAoverBAS shell help"
  printf "%s\n" "  <basmodule> = filename of project source .xba to copy (e.g. Module1)"
  printf "%s\n" ""
  printf "%s\n" "Results: /Basic/<basmodule/<basmodule>.xba updated from Import/<basmodule>.xba"
  printf "%s\n" ""
  exit 0
 else
  printf "%s" "CopyXBAoverBAS.sh -h|<basmodule>"
  printf "%s\n" " unrecognized option '$P1' - exiting."
  exit 1
 fi		# have -h
fi	# have -
if [ -z "$P1" ];then
 echo "CopyXBAoverBAS <xbamodule> missing parameter(s) - abandoned."
 read -p "Enter ctrl-c to remain in Terminal: "
 exit 1
fi
cd $srcpath
if ! test -f $P1.xba;then
 echo "CopyXBAoverBAS $srcpath/$P1.xba not found for copy - abandoned."
 read -p "Enter ctrl-c to remain in Terminal: "
 exit 1
fi
if [ $srcpath/$P1.xba -nt $targpath/$P1.xba ];then
 cp -u $srcpath/$P1.xba $targpath/$P1.xba
 echo "$P1.xba updated in /Basic folder."
 ~/sysprocs/LOGMSG "  CopyXBAoverBAS - $P1.xba been updated in /Basic folder from Import."
else
 echo "$srdpath/$P1.xba copy skipped - /Basic file is not older.**"
 ~/sysprocs/LOGMSG "**$ CopyXBAoverBAS $srcpath/$P1.xba copy skipped - /Basic file is not older.**"
fi
# end CopyXBAoverBAS.

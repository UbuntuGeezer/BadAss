#!/bin/bash
# GetXBAModule - Copy BadAss/<xbamodule>.xba mod to Import folder.
#	4/19/25.	wmk.
#
# Usage. bash  GetXBAModule.sh -h|<xbamodule>
#
#	-h = only display GetXBAModule shell help
#	<xbamodule> = module name from library (e.g. Module1)
#
# Exit.	*gitbase/<xbamodule>.xba copied to *gitbase/Import
#
# Modification History.
# ---------------------
# 4/19/25.	wmk.	(automated) Modification History sorted.
# 4/19/25.	wmk.	-h option support. 
# 1/12/24.	wmk.	code checked for Accounting/BadAss compatibility. 
# 11/12/23.	wmk.	(automated) Version 3.0.9 *libpath introduced. 
# 11/12/23.	wmk.	(automated) Version 3.0.6 Make old paths removed. 
# 9/17/23.	wmk.	*folderbase corrected for HPPavilion; /src added to 
# 9/17/23.	 target path.
# 8/23/23.	wmk.	ver2.0 change to use /src folder; bug fix *github 
# 8/23/23.	 corrected to *gitbase. 
# 8/22/23.	wmk.	edited for FLsara86777. 
# 8/13/23.	wmk.	edited for MNcrwg44586; *folderbase conditional. 
# 6/24/23.	wmk.	modified for FLsara86777. 
# 3/8/22.	wmk.	exit added if file not found error. 
# 3/7/22.	wmk.	original code. 
#
# P1=<xbamodule>
#
projbase=$libbase/src/Projects-Geany/EditBas
gitbase=$libbase
P1=$1
# -h option code
if [ "${P1:0:1}" == "-" ];then
 option=${P1,,}
 if [ "$option" == "-h" ];then
  printf "%s\n" "GetXBAModule - Copy BadAss/<xbamodule>.xba mod to Import folder."
  printf "%s\n" "GetXBAModule.sh -h|<xbamodule>"
  printf "%s\n" ""
  printf "%s\n" "  -h = only display GetXBAModule shell help"
  printf "%s\n" "  <xbamodule> = module name from library (e.g. Module1)"
  printf "%s\n" ""
  printf "%s\n" "Results: *gitbase/<xbamodule>.xba copied to *gitbase/Import"
  printf "%s\n" ""
  exit 0
 else
  printf "%s" "GetXBAModule.sh -h|<xbamodule>"
  printf "%s\n" " unrecognized option '$P1' - exiting."
  exit 1
 fi		# have -h
fi	# have -
if [ -z "$P1" ];then
 echo "GetXBAModule -h|<xbamodule> missing parameter(s) - abandoned.**"
 read -p "Enter ctrl-c to remain in Terminal: "
 exit 1
fi
cd $gitbase
if ! test -f $gitbase/$P1.xba;then
 echo "** file GitHub/$P1.xba not found for copy - GetXBAModule abandoned.**"
 read -p "Enter ctrl-c to remain in Terminal: "
 exit 1
fi
if [ $gitbase/$P1.xba -ot $gitbase/src/Import/$P1.xba ];then
 echo "**$P1.xba copy skipped - Import/$P1.xba is newer.**"
 ~/sysprocs/LOGMSG "**$ GetXBAModule $P1.xba copy skipped - Import/$P1.xba is newer.**"
else
 cp -uv $gitbase/$P1.xba $gitbase/src/Import/$P1.xba
 echo "$P1.xba copied to Import folder."
 ~/sysprocs/LOGMSG "  GetXBAModule $P1.xba copied to Import folder."
fi
# end GetXBAModule.sh.

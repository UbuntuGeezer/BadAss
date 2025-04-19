#!/bin/bash
# GetXBAModule - Copy /Libraries-Project/FLsara86777 .xba module to Import folder.
#	1/12/24.	wmk.
#
# Usage. bash  GetXBAModule.sh  <xbamodule>
#
#	<xbamodule> = module name from library (e.g. Module1)
#
# Exit.	*gitbase/<xbamodule>.xba copied to *gitbase/Import
#
# Modification History.
# ---------------------
# 8/22/23.	wmk.	edited for FLsara86777.
# 8/23/23.	wmk.	ver2.0 change to use /src folder; bug fix *github
#			 corrected to *gitbase.
# 9/17/23.	wmk.	*folderbase corrected for HPPavilion; /src added to
#			 target path.
# 11/12/23.	wmk.	(automated) Version 3.0.6 Make old paths removed.
# 11/12/23.	wmk.	(automated) Version 3.0.9 *libpath introduced.
# 1/12/24.	wmk.	code checked for Accounting/BadAss compatibility.
# Legacy mods.
# 8/13/23.	wmk.	edited for MNcrwg44586; *folderbase conditional.
# Legacy mods.
# 6/24/23.	wmk.	modified for FLsara86777.
# Legacy mods.
# 3/7/22.	wmk.	original code.
# 3/8/22.	wmk.	exit added if file not found error.
projbase=$libbase/src/Projects-Geany/EditBas
gitbase=$libbase
P1=$1
if [ -z "$P1" ];then
 echo "GetXBAModule <xbamodule> missing parameter(s) - abandoned.**"
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

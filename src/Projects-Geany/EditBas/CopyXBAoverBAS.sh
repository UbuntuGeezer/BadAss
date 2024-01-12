#!/bin/bash
# 2023-11-12.	wmk.	(automated) Version 3.0.6 paths eliminated (HPPavilion).
# 2023-11-12.	wmk.	(automated) Version 3.0.9 *libpath introduced (HPPavilion).
# CopyXBAoverBAS - Copy FLsara86777 Import module .xba code over to Basic folder.
#	11/13/23.	wmk.
#
# Usage.	bash CopyXBAoverBAS.sh <basmodule>
#
#	<basmodule> = filename of project source .xba to copy (e.g. Module1)
#
# Exit.	/Basic/<basmodule/<basmodule>.xba updated from Import/<basmodule>.xba
#
# Modification History.
# ---------------------
# 11/12/23.	wmk.	(automated) Version 3.0.6 Make old paths removed.
# 11/12/23.	wmk.	(automated) Version 3.0.9 *libpath introduced.
# 11/13/23.	wmk.	*srcpath corrected.
# Legacy mods.
# 8/22/23.	wmk.	edited for FLsara86777; *codebase used throughout.
# 8/23/23.	wmk.	ver2.0 change to use /src as parent folder.
# 9/17/23.	wmk.	Exit path corrected.
# Legacy mods.
# 8/13/23.	wmk.	*pathbase, *codebase unconditional; corrected to point to
#			 GitHub/Libraries-Project/MNcrwg44586.
# Legacy mods.
# 6/25/23.	wmk.	original code; adpated from Territories.
# Legacy mods.
# 3/7/22.	wmk.	original code.
# 3/8/22	wmk.	description clarified.
# 4/24/22.	wmk.	*pathbase* env var included.
# 9/23/22.  wmk.    (automated) CB *codebase env var support.
P1=$1
srcpath=$libbase/src/Import
targpath=$libbase/src/Basic/$P1
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

#!/bin/bash
# DoSed.sh - sed editing for BadAssLibrary project.
#	8/23/22.	wmk.
#
# Usage.	bash DoSed.sh  <module-name>
#
#	<module-name> = target name for .xba (e.g. Module1)
#
# Exit.	awkBuild.tmp -> awkBuild.txt
#	MakeBadAssLibrary.tmp -> MakeBadAssLibrary
#	XBAHeader.tmp -> XBAHeader.txt
#
# Modification History.
# ---------------------
# 8/23/22.	wmk.	original code.
P1=$1
if [ -z "$P1" ];then
 echo "DoSed <module-name> missing parameter(s) - abandoned."
 exit 1
fi
# if called from Build menu, env vars not set.
if [ -z "$folderbase" ];then
 if [ "$USER" == "ubuntu" ];then
  export folderbase=/media/ubuntu/Windows/Users/Bill
 else
  export folderbase=$HOME
 fi
fi
if [ -z "$pathbase" ];then
 export pathbase=$folderbase/Accounting
fi
sed "s?<module-name>?$P1?g" \
  awkBuild.tmp > awkBuild.txt
sed "s?<module-name>?$P1?g" \
  MakeBadAssLibrary.tmp > MakeBadAssLibrary
sed "s?<module-name>?$P1?g" \
  XBAHeader.tmp > XBAHeader.txt
# end DoSed.sh

#!/bin/bash
echo " ** CopyToGit.sh out-of-date **";exit 1
# CopyToGit.sh - sed editing for BadAssLibrary project.
#	8/25/22.	wmk.
#
# Usage.	bash CopyToGit.sh  <module-name>
#
#	<module-name> = Module to copy (e.g. Module1)
#
# Entry.	*WINGIT_PATH = base path for GitHub projects.
#
# Exit.	<module-name>.xba  - > *WINGIT_PATH/Libraries-Project/BadAss
#
#	MakeBadAssLibrary.tmp -> MakeBadAssLibrary
#	XBAHeader.tmp -> XBAHeader.txt
#
# Modification History.
# ---------------------
# 8/25/22.	wmk.	original code.
P1=$1
if [ -z "$P1" ];then
 echo "CopyToGit <module-name> missing parameters(s) - abandoned."
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
cp -pv $P1.xba $WINGIT_PATH/Libraries-Project/BadAss
echo "CopyToGit complete."
# end CopyToGit.sh

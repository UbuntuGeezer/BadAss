#!/bin/bash
echo " ** UpdateBasicFiles.sh out-of-date **";exit 1
# Note. UpdateBasicFiles is a leftover from the original EditBas project
# where the .bas files were edited within the EditBas project folder. Now
# the .bas files reside in the Basic/Module(n) folders, so this shell is
# no longer necessary.
#
# UpdateBasicFiles - Copy newer .bas files from project to /Basic folder.
#	4/20/25.	wmk.
#
# Usage. bash   UpdateBasicFiles.sh [-h]
#
#	-h = (optional) only display <filename> shell help
#
# Modification History.
# ---------------------
# 4/20/25.	wmk.	-h option support.
# 11/13/23.	wmk.	(automated) Version 3.0.6 Make old paths removed.
# 11/13/23.	wmk.	(automated) Version 3.0.9 *libpath introduced.
# 3/8/22.	wmk.	original code; adapted from PutXBAModule.
#
projbase=$codebase/src/Projects-Geany/EditBas
gitbase=$codebase
echo "**WARNING - all .bas files in EditBas/ folder will be copied over older files"
read -p "  in /Basic folder... Do you wish to proceed (Y/N)? "
yn=${REPLY,,}
if [ "$yn" == "n" ];then
 echo "UpdateBasicFiles abandoned."
 exit
else
 cp -uv $projbase/*.bas $folderbase/Territories/Basic 
 echo "EditBas/*.bas copied to /Basic folder."
 ~/sysprocs/LOGMSG "  UpdateBasicFiles - EditBas/*.bas copied to /Basic folder."
fi
read -p "Do you wish to remove all .bas files from /EditBas (y/n)?"
yn=${REPLY,,}
if [ "$yn" == "y" ];then
 rm *.bas;rm *.ba1
 echo " *.bas files removed from /EditBas."
 ~/sysprocs/LOGMSG " UpdateBasicFiles - *.bas files removed from /EditBas."
else
 echo " *.bas files in /EditBas retained."
 ~/sysprocs/LOGMSG " UpdateBasicFiles - *.bas files in /EditBas retained."
fi
# end UpdateBasicFiles.sh.



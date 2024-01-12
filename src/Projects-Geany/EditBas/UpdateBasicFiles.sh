#!/bin/bash
echo " ** KillShell.sh out-of-date **";exit 1
# 2023-11-13.	wmk.	(automated) Version 3.0.6 paths eliminated (HPPavilion).
# 2023-11-13.	wmk.	(automated) Version 3.0.9 *libpath introduced (HPPavilion).
# UpdateBasicFiles - Copy newer .bas files from project to /Basic folder.
#	3/8/22.	wmk.
#
# Modification History.
# ---------------------
# 11/13/23.	wmk.	(automated) Version 3.0.6 Make old paths removed.
# 11/13/23.	wmk.	(automated) Version 3.0.9 *libpath introduced.
# 3/8/22.	wmk.	original code; adapted from PutXBAModule.
projbase=$folderbase/Territories/Projects-Geany/EditBas
gitbase=$folderbase/GitHub/Libraries-Project/Territories
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



#!/bin/bash
echo " ** KillShell.sh out-of-date **";exit 1
# 2023-11-13.	wmk.	(automated) Version 3.0.6 paths eliminated (HPPavilion).
# 2023-11-13.	wmk.	(automated) Version 3.0.9 *libpath introduced (HPPavilion).
# UpdateLoadSource.sh - Update Calc library load source.
# 6/21/23.	wmk.
#
# Usage. bash  UpdateLoadSource.sh  <libname>
#
#	<libname> = library name (e.g. Territories)
#
# Entry. Libraries-Project/<libname>/Release = update for library
#
# Exit. Libraries-Project/<libname> = updated with newer files from
#	<libname>/Release (cp -uv utility)
#
# Modification History.
# ---------------------
# 11/13/23.	wmk.	(automated) Version 3.0.6 Make old paths removed.
# 11/13/23.	wmk.	(automated) Version 3.0.9 *libpath introduced.
# 6/21/23.	wmk.	original shell (template)
#
# Notes. 
#
# P1=<libname>
#
P1=$1
if [ -z "$P1" ];then
 echo "UpdateLoadSource <libname> missing parameter(s) - abandoned."
 read -p"Enter ctrl-c to remain in Terminal: "
 exit 1
fi
#	Environment vars:
if [ -z "$TODAY" ];then
 . ~/GitHub/TerritoriesCB/Procs-Dev/SetToday.sh
#TODAY=2022-04-22
fi
#procbodyhere
pushd ./ > /dev/null
echo " ** You are about to overwrite the running source for $P1..."
read -p " Do you wish to continue (y/n)? "
yn=${REPLY^^}
if [ "$yn" != "Y" ];then
 echo "UpdateLoadSource terminated by user; $P1 unchanged."
 read -p "Enter ctrl-c to remain in Terminal: "
 exit 1
fi
cd $pathbase/Release
cp -uv * $pathbase
popd > /dev/null
#endprocbody
echo "  UpdateLoadSource complete."
~/sysprocs/LOGMSG "  UpdateLoadSource complete."
# end UpdateLoadSource.sh

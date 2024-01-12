#!/bin/bash
# GenFolders.sh - Create and initialize BadAss/Basic folders.
#	1/12/24.	wmk.
#
# Usage. bash  GenFolders.sh
#
# Entry. 
#
# Dependencies. GetBasList.sh (for generating <xbamodule>Bas.txt listing of all
#	.bas macro blocks within each <xbamodule>.xba. This file will be used by
#	EditBas to extract the source for all .bas macros into Basic/Module_x.
#
# Modification History.
# ---------------------
# 1/12/24.	wmk.	Modified for BadAss/Basic folders.
# Legacy mods.
# 11/13/23.	wmk.	(automated) Version 3.0.6 Make old paths removed.
# 11/13/23.	wmk.	(automated) Version 3.0.9 *libpath introduced.
# 6/22/23.	wmk.	original shell.
# 6/25/23.	wmk.	skip *mkdir if module folder already exists.
# 8/21/23.	wmk.	switch to *codebase for build; *codebase, *pathbase defs
#			 unconditional; documentation improved.

#	Environment vars:
#if [ -z "$TODAY" ];then
# . ~/GitHub/TerritoriesCB/Procs-Dev/SetToday.sh
#TODAY=2022-04-22
#fi
#procbodyhere
pushd ./ > /dev/null
# ensure folders Procs-Dev, and Basic/Import, Basic/Release defined for
# FLsara86777.
# get list of .xba files (modules) in FLsara86777 main folder.
# ensure folder 'Basic' exists.
# define folders under 'Basic' using list of .xba files.
# copy .xba files to their matching folders
# run GetBasList.sh on each module.
libbase=$codebase/BadAss
cd $libbase/src
if ! test -d Procs-Dev;then
 mkdir Procs-Dev
fi
if ! test -d Basic;then
 mkdir Basic
fi
cd Basic
if ! test -d Import;then
 mkdir Import
fi
if ! test -d Release;then
 mkdir Release
fi
ls *.xba > $TEMP_PATH/XBAList.txt
mawk -F "." '{print $1}' $TEMP_PATH/XBAList.txt > $TEMP_PATH/ModuleList.txt
# loop creating Basic/Modulex folders.
file=$TEMP_PATH/ModuleList.txt
cd Basic
while read -e;do
 if ! test  -d $REPLY;then
  mkdir $REPLY
 else
  echo "  folder Basic/$REPLY already exists - skipping mkdir.."
 fi
done < $file
# loop initilialing .xba files.
popd > /dev/null
file=$TEMP_PATH/XBAList.txt
while read -e;do
 fname=$REPLY
 echo "$fname" > fname.txt
 mawk -F "." '{print "cp -pv " $1 ".xba Basic/" $1}' fname.txt > cpyfile.sh
 chmod +x cpyfile.sh
 ./cpyfile.sh
done < $file
# loop initalizing <module>Bas.txt files.
projpath=$libbase/src/Projects-Geany/LibBuild
pushd ./ > /dev/null
cd $projpath
file=$TEMP_PATH/ModuleList.txt
while read -e;do
 modname=$REPLY
 ./GetBasList.sh $modname
done < $file
# now copy dialogs to Dialogs folder
cd $libbase/src/Basic
if ! test -d Dialogs;then
 mkdir Dialogs
fi
cd $libbase 
ls *.xdl > $TEMP_PATH/XDLList.txt
mawk -F "." '{print $1}' $TEMP_PATH/XBAList.txt > $TEMP_PATH/DialogList.txt
if test -s $TEMP_PATH/DialogList.txt;then
 cp -pv *.xdl $libbase/src/Basic/Dialogs
fi
#endprocbody
echo "  GenFolders complete."
~/sysprocs/LOGMSG "  GenFolders complete."
echo ""
echo "*** Now Use EditBas project to perform module code maintenance."
echo "*** Be sure to run GetModuleXBAHdr to extract XML header."
read -p "Enter ctrl-c to remain in Terminal: "
exit 0
# end GenFolders.sh

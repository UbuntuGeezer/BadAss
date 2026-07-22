# folders.sh - function definitions for (Windows) Accounting/BadAss folders. 1/12/24. wmk.
#	7/22/26.	wmk.
#
# Modification History.
# ---------------------
# 7/22/26.  wmk.    modified for Windows 11; paths changed for Windows.
# 4/20/25.	wmk.	cdv added to change to version control folder.
# 10/23/23.	wmk.	paths for Lenovo/Accounting folders.
# 1/4/24.	wmk.	cdb added to change to Basic folders.
# 1/5/24.	wmk.	cda mod to accept P1 subfolder; cdb mod using *acct_yr.
# 1/12/24.	wmk.	cda cdb, cdp, cdj, cdp for BadAss library project.
# Legacy mods.
# 9/4/23.	wmk.	corrections missed by (automated) corrections.
# 9/11/23.	wmk.	cdt, cds accept subfolders
# 9/13/23.	wmk.	ver2.0 paths to FLsara86777 for codebase.
# 9/14/23.	wmk.	"huh" updated; cdpb (change to *pathbase) added.
# 9/21/23.	wmk.	ver 2.0.2 merge with HPPavilion2 changes.
function cda(){
 P1=$1
 cd C:/Users/vncwm/linux/BadAss/$P1
}
function cdab(){
 cd C:/Users/vncwm/linux/BadAss/src/Projects-Geany/ArchivingBackups
}
function cdb(){
 P1=$1
 cd C:/Users/vncwm/linux/BadAss/src/Basic/$P1
}
function cdc(){
 P1=$1
 cd C:/Users/vncwm/linux/BadAss/$P1
}
function cdd(){
 echo "cdd stubbed."
}
function cdg(){
 P1=$1
 cd C:/Users/vncwm/linux/$P1
}
function cdj(){
 P1=$1
 cd C:/Users/vncwm/linux/BadAss/src/Projects-Geany/$P1
}
function cdp(){
 cd C:/Users/vncwm/linux/BadAss/src/Procs-Dev
}
function cdr(){
 cd C:/Users/vncwm/linux/BadAss/src/Release
}
function cdt(){
 echo "cdt stubbed."
}
function cdv(){
 cd $folderbase/Accounting/AcctngVersion
}
function cdts(){
 echo "cdts stubbed."
}
function cds(){
 echo "cds stubbed."
}
function cdss(){
 echo "cdss stubbed."
}
function huh(){
 echo "/BadAss folders.sh functions:"
 echo "cda - change to BadAss/ folder."
 echo "cdab - change to BadAss/../ArchivingBackups project folder."
 echo "cdb - change to BadAss/Basic folder."
 echo "cdc - change to BadAss/*P1"
 echo "cdd - stubbed"
 echo "cdg - change to ~/linux."
 echo "cdj - change to BadAss/../Projects-Geany/*P1 project folder."
 echo "cdp - change to BadAss/../Procs-Dev."
 echo "cdt - stubbed."
 echo "cdts - stubbed."
 echo "cds - stubbed."
 echo "cdss - stubbed."
 echo "help/huh - list this list to terminal."
}
function help(){
 huh
 }

# folders.sh - function definitions for Accounting/BadAss folders. 1/12/24. wmk.
# Modification History.
# ---------------------
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
 cd $folderbase/Accounting/BadAss/$P1
}
function cdab(){
 cd $folderbase/Accounting/Projects-Geany/ArchivingBackups
}
function cdb(){
 P1=$1
 cd $folderbase/Accounting/BadAss/src/Basic/$P1
}
function cdc(){
 P1=$1
 cd $folderbase/Accounting/$P1
}
function cdd(){
 echo "cdd stubbed."
}
function cdg(){
 P1=$1
 cd $folderbase/GitHub/$P1
}
function cdj(){
 P1=$1
 cd $folderbase/Accounting/BadAss/src/Projects-Geany/$P1
}
function cdp(){
 cd $folderbase/Accounting/BadAss/src/Procs-Dev
}
function cdr(){
 cd $folderbase/Accounting/BadAss/src/Release
}
function cdt(){
 echo "cdt stubbed."
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
 echo "Accounting/BadAss folders.sh functions:"
 echo "cda - change to Accounting/ folder."
 echo "cdab - change to Accounting/../ArchivingBackups project folder."
 echo "cdb - change to Accounting/$acct_yr/Basic folder."
 echo "cdc - change to Accounting/*P1"
 echo "cdd - stubbed"
 echo "cdg - change to *HOME/GitHub."
 echo "cdj - change to Accounting/../Projects-Geany/*P1 project folder."
 echo "cdp - change to Accounting/../Procs-Dev."
 echo "cdt - stubbed."
 echo "cdts - stubbed."
 echo "cds - stubbed."
 echo "cdss - stubbed."
 echo "help/huh - list this list to terminal."
}
function help(){
 huh
 }

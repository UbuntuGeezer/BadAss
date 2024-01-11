# GetBasList.sh - Get list of all .bas modules.
#	8/23/22.	wmk.
#
# Usage. bash  GetBasList.sh <module-name>
#
#	<module-name> = module for which to generate BasList.txt
#
# Exit. BasList.txt has list of .bas filenames preceded by <module-name>/
#
P1=$1
if [ -z "$P1" ];then
 echo "GetBasList <module-name> missing parameter - abandoned."
 exit 1
fi
ls $P1/*.bas > BasList.txt

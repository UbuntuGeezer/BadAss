#!/bin/bash
# SetToday.sh - set *TODAY and *TODAY1 environment vars.
#	1/12/24.	wmk.
#
# Usage. . SetToday.sh [-v]
#
#	-v = (optional) verbose - printf "%s\n" *TODAY1 for user verification.
#
# Exit.	*TODAY = today's date yyyy-mm-dd
#	*TODAY1 = today's date mm/dd/yy
#
# Modification History.
# ---------------------
# 1/12/24.	wmk.	integrated into BadAss/src/Procs-Dev.
# Legacy mods.
# 11/3/23.	wmk.	original code.
# 11/15/23.	wmk.	(automated) Version 3.0.6 paths eliminated (HPPavilion).
# 12/1/23.	wmk.	-v option added.
#
# Notes.. only works if year >=10.
# [P1=-v]
P1=${1,,}
if [ ! -z "$P1" ];then
 if [ "$P1" != "-v" ];then
  printf "%s\n" "SetToday [-v] unrecognized '$P1' - abandoned."
  exit 1
 fi
fi
date +%Y-%m-%d|mawk 'BEGIN {print "#!/bin/bash"}{print "export TODAY=" $0}' > $TEMP_PATH/SetToday.sh
date +%m/%d/%y|mawk 'BEGIN {print "#!/bin/bash"}{print "export TODAY1=" $0}' >> $TEMP_PATH/SetToday.sh
sed -i  '2s?^0??1;s?/0?/?g' $TEMP_PATH/SetToday.sh
chmod +x $TEMP_PATH/SetToday.sh
. $TEMP_PATH/SetToday.sh
if [ ! -z "$P1" ];then printf "%s\n" $TODAY1;fi
# end SetToday.

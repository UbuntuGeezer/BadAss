function doit(){
 cp -p $P1 $TEMP_PATH
 chmod +x $TEMP_PATH/$P1
 . $TEMP_PATH/$P1 $A1
}

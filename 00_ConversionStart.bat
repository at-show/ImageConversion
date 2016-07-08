echo ########## start ########## [%date% %time%] >> 02_Conversion.log
cscript 01_Conversion.vbs >> 02_Conversion.log 2>&1
echo ##########  end  ########## [%date% %time%] >> 02_Conversion.log

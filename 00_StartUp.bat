echo ########## start ########## [%date% %time%] >> 03_Conversion.log
cscript 01_Conversion.vbs >> 03_Conversion.log 2>&1
echo ##########  end  ########## [%date% %time%] >> 03_Conversion.log

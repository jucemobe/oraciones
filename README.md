# oraciones
-automatizacion para descargar las oraciones de la liturgia y lecturas diariamente
C:\tmp\testigo>C:\Users\"Julio Cesar"\Downloads\WinWGetPortable\App\WinWget\wget\wget.exe "http://www.eltestigofiel.org/index.php?idu=lt_2134&id_fecha=1-03-2020&idd=69&hora=1" -O "C:\tmp\testigo\2020\03\1\laudes.txt" -F
C:\tmp\testigo>cscript //Nologo delete.vbs C:\tmp\testigo\2020\03\1\laudes.txt  
C:\tmp\testigo\2020\03\1>cscript //Nologo ejemplo.vbs C:\tmp\testigo\2020\03\1\laudes.xml

****29mar2020
C:\tmp\testigo>C:\Users\"Julio Cesar"\Downloads\WinWGetPortable\App\WinWget\wget\wget.exe "https://www.eltestigofiel.org/index.php?idu=lt_liturgia&dia=29&mes=3&ano=2020"  -O "C:\tmp\testigo\2020\03\29\total.txt" -F 
C:\tmp\testigo>C:\Users\"Julio Cesar"\Downloads\WinWGetPortable\App\WinWget\wget\wget.exe "http://www.eltestigofiel.org/index.php?idu=lt_2134&id_fecha=29-3-2020&idd=809&hora=1" -O "C:\tmp\testigo\2020\03\29\laudes.txt" -F 
C:\tmp\testigo>cscript //Nologo delete.vbs C:\tmp\testigo\2020\03\29\laudes.txt "Acudamos a nuestro Redentor, que nos concede estos" "Digamos ahora, todos juntos, la " true
C:\tmp\testigo\2020\03\29>cscript //Nologo ejemplo.vbs C:\tmp\testigo\2020\03\29\laudes.xml
C:\tmp\testigo>C:\Users\"Julio Cesar"\Downloads\WinWGetPortable\App\WinWget\wget\wget.exe "http://www.eltestigofiel.org/index.php?idu=lt_2134&id_fecha=29-3-2020&idd=119&hora=1" -O "C:\tmp\testigo\2020\03\29\laudes.txt" -F 
C:\tmp\testigo>cscript //Nologo delete.vbs C:\tmp\testigo\2020\03\29\laudes.txt "Los que celebramos  hoy el principio de nuestra " "a nuestro Padre con la oraci" false
C:\tmp\testigo\2020\03\29>cscript //Nologo ejemplo.vbs C:\tmp\testigo\2020\03\29\laudes.xml

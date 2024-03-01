REM @echo off
chcp 65001 
"C:\Python373\python.exe" "P:\h978\insect\Kahanpää\SMALL_TOOLS\gbif_taxon_fetch\gbiftaxonfetch.py" %*
set /p id="Enter to quit"
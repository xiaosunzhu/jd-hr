@echo off
rmdir /S/Q punch-check
python update-setup.py py2exe
python punch-setup.py py2exe
rename dist punch-check
xcopy config punch-check\config\
rem java -jar ..\tools\Markdown2HTML-1.0-SNAPSHOT.jar docs\参考手册.md -out punch-check\参考手册.html
rem java -jar ..\tools\Markdown2HTML-1.0-SNAPSHOT.jar docs\使用说明.md -out punch-check\使用说明.html
copy docs\参考手册.html punch-check\
copy docs\使用说明.html punch-check\
copy docs\changelist.txt punch-check\
pause
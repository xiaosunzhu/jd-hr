@echo off
rmdir /S/Q punch-check
python update-setup.py py2exe
python punch-setup.py py2exe
rename dist punch-check
xcopy config punch-check\config\
rem java -jar ..\tools\Markdown2HTML-1.0-SNAPSHOT.jar docs\�ο��ֲ�.md -out punch-check\�ο��ֲ�.html
rem java -jar ..\tools\Markdown2HTML-1.0-SNAPSHOT.jar docs\ʹ��˵��.md -out punch-check\ʹ��˵��.html
copy docs\�ο��ֲ�.html punch-check\
copy docs\ʹ��˵��.html punch-check\
copy docs\changelist.txt punch-check\
pause
echo LookOut > ..\readme.md
FINDSTR  /V /C:"LookOut{#mainpage}"  readme.md >> ..\readme.md
del readme.md~
del LookOut.cls~


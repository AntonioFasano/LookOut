::Run before commit to Github
::Cleanup and creation of Github readme

 
echo LookOut > ..\readme.md
FINDSTR  /V /C:"LookOut{#mainpage}"  readme.md >> ..\readme.md

del *.cfg~
del *.cls~
del *.cmd~
del *.js~
del *.md~

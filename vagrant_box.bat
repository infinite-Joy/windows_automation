@echo off
REM this can be fired up by just double-clicking

REM go to vagrant folder
cd /d "I:\vagrant"

REM fire up vagrant
vagrant up

REM fire up putty
REM usage putty username@servername -pw password -P port_number
REM just type exit in the putty and this will exit
putty.exe vagrant@127.0.0.1 -pw vagrant -P 2222

REM stop the vm
putty.exe -ssh "127.0.0.1" -l vagrant -pw vagrant -m "G:\Python\projects\pywinauto\windows_automation\shutdown_command.txt"
vagrant halt

pause
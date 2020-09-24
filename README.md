<div align="center">

## AA Very Easy Shutdown Methods for Novice Users


</div>

### Description

Allows a novice user to shutdown/restart/log off/force shutdown with or without a specific time length his/her computer with a simple one line of code
 
### More Info
 


<span>             |<span>
---                |---
**Submitted On**   |
**By**             |[ukWaqas](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByAuthor/ukwaqas.md)
**Level**          |Beginner
**User Rating**    |4.6 (23 globes from 5 users)
**Compatibility**  |VB 4\.0 \(32\-bit\), VB 5\.0, VB 6\.0, VBA MS Access, VBA MS Excel
**Category**       |[Miscellaneous](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByCategory/miscellaneous__1-1.md)
**World**          |[Visual Basic](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByWorld/visual-basic.md)
**Archive File**   |[](https://github.com/Planet-Source-Code/ukwaqas-aa-very-easy-shutdown-methods-for-novice-users__1-63282/archive/master.zip)





### Source Code

<b> Shutting down the computer in Visual Basic </b>
There are different commands you can give to shut down the computer.
Type thiese in any procedure e.g. Command1.Click() <br><br>
<b>To Log off: </b> <br>
 retval=shell("RUNDLL32 SHELL32.DLL,SHExitWindowsEx 0",1) <br> <br>
<b>To Shut Down: </b> <br>
 retval=shell("RUNDLL32 SHELL32.DLL,SHExitWindowsEx 1",1) <br> <br>
<b>To Reboot: </b>
 retval=shell("RUNDLL32 SHELL32.DLL,SHExitWindowsEx 2",1) <br> <br>
<b>To Force ShutDown: </b>
 retval=shell("RUNDLL32 SHELL32.DLL,SHExitWindowsEx 4") <br> <br>
The above code can also be inserted into timer so the computer shutdowns after a period of time. <br>
However, if you are a complete novice and use Windows Xp, then use the following code: <br> <br>
retval=shell("SHUTDOWN /L /T:15 /Y /C",1) <br> <br>
The above code will shutdown in 15 seconds, and close all existing programs without saving. <br>
If however you wish to stop the shutdown during these 15 seconds, then a simple ccommand: <br> <br>
retval=shell("shutdown /a",1) <br> <br>
should suffice. Note that the 15 seconds can be adjusted to seconds of your choice. Insert /n instead of /y for prompt to save exisiting applications and remove the final parameter /c if you do not want shutdown to be cancelled.
Hope the above Article helps you in understanding in different ways to shutdown. <br>
Please, please, please vote for me! It is my first submission


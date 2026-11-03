# Local VB6 installation and compiler

## Dependency boundary

The code must use as little dependencies as possible. The program is portable, and must remain so.


The VB6 installation and compiler are located at `C:\Program Files (x86)\Microsoft Visual Studio\VB98`.
Agents are authorized to run the compiler to compile and test as needed.

## VB6 compiler
VB6 projects are compiled by the IDE, `VB6.exe`.

## VB6 command-line arguments
VB6[.EXE] [{/run | /r}] | [/runexit] | [{/make | /m}] projectname] [/out filename] [/outdir path] [/d const=value[{:constN=valueN}]] [/mdi | /sdi] [{/cmd argument | /c argument}]

| Argument | Behavior |
|---|---|
| /run or /r projectname | Tells Visual Basic to compile projectname and run it, using the arguments stored in the Command Line Arguments field of the Make tab of the Project Properties dialog box. |
| /runexit projectname | Tells Visual Basic to compile projectname and run it. Visual Basic will exit when the project returns to design mode. |
| /make or /m projectname | Tells Visual Basic to compile projectname and make an executable file from it, using the existing settings stored in the project file. |
| /out filename | Specifies a file to receive errors when you build using /m or /runexit. If you do not use /out, command line build errors are displayed in a message box. |
| /outdir path | Specifies a directory path to place all output files in when using /make. |
| /cmd or /c argument | Puts argument in the Command Line Arguments field in the Make tab of the Project Properties dialog box. When used, this must be the last switch on the command line. |
| /mdi or /sdi | Changes the Visual Basic environment to either Single Document Interface (SDI) or Multiple Document Interface (MDI) mode. Visual Basic remains in this mode until you change it. |

You may use a group name in place of projectname in any of the above switches.
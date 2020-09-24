<div align="center">

## CommConfigDialog


</div>

### Description

Wrapper for CommConfigDialog API call, makes it very easy to show the system configuration dialog for serial ports. Just set the CommPort property, call the ShowCommConfig method and the dialog is displayed. After the user clicks OK, simply read the Settings property and paste this into the MSComm.Settings property. No hassle with long API declarations or tedious settings of system variables...
 
### More Info
 
Set CommPort property to the same value as your MSComm control and execute the ShowCommConfig method.

Compile the control, copy it to your system directory, register it into Windows with "regsvr32 edldlg32.ocx", then you can use it in the VB IDE (Components tab, click "Add" and then check "CommConfigDialog"). Paste it onto your form together with an MSComm control and you're set!

Read Settings property as follows:

MSComm1.Settings=CommComfigDialog1.Settings

Also has single properties for baud rate, databits, handshake, parity, stop bits.

None that I'm aware of...


<span>             |<span>
---                |---
**Submitted On**   |2002-08-24 00:45:18
**By**             |[Wolfgang Lindke](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByAuthor/wolfgang-lindke.md)
**Level**          |Intermediate
**User Rating**    |5.0 (20 globes from 4 users)
**Compatibility**  |VB 5\.0, VB 6\.0
**Category**       |[Custom Controls/ Forms/  Menus](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByCategory/custom-controls-forms-menus__1-4.md)
**World**          |[Visual Basic](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByWorld/visual-basic.md)
**Archive File**   |[CommConfig1218018232002\.zip](https://github.com/Planet-Source-Code/wolfgang-lindke-commconfigdialog__1-38229/archive/master.zip)

### API Declarations

Type DCB, Type COMMCONFIG, API call CommConfigDialog






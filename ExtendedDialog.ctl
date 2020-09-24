VERSION 5.00
Begin VB.UserControl CommDialog 
   Appearance      =   0  'Flat
   BorderStyle     =   1  'Fixed Single
   CanGetFocus     =   0   'False
   ClientHeight    =   435
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   435
   InvisibleAtRuntime=   -1  'True
   ScaleHeight     =   420
   ScaleMode       =   0  'User
   ScaleWidth      =   406
   ToolboxBitmap   =   "ExtendedDialog.ctx":0000
   Begin VB.Image Image1 
      Height          =   420
      Left            =   0
      Picture         =   "ExtendedDialog.ctx":0312
      Top             =   0
      Width           =   420
   End
End
Attribute VB_Name = "CommDialog"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True

'***************************************************************************************************
'
' This code and the user control is (C) 1998-2002 by Wolfgang Lindke (network.admin@hispeed.ch)
' You may use the code at your own risk and redistribute it freely as long as this copyright message
' remains intact. Please feel free to send comments or ideas for enhancements to my email address.
' No warranties are given, neither implied nor expressed!
'
'
' This control is a wrapper for the CommConfigDialog API call
' The main reason for me to write this control was to have an easy and fast method of
' displaying a serial port configuration dialog and get the selected settings back into my
' main program. This also provides a better readability of the main code (No bloody
' declares of DCB and COMMCONFIG structures and no API call definitions).
'
' Under 9x/ME this was also possible with a ConfigurePort call but under NT/W2K this works only
' for printers. On NT/W2K/XP the CommConfigDialog API call is the only way to display the system
' configuration dialog for serial ports. But I wanted a generic solution for all operating systems
' without tedious GetVersionEx API calls...
'
' Properties:
' Name          Access      Data Type   Description
' ----          ------      ---------   -----------
' CommPort      Read/Write  Integer     Sets/gets the CommPort that we want to configure
' Baudrate      Read only   Long        Gets the selected baud rate from the dialog
' Databits      Read only   Integer     Gets the amount of databits selected in the configuration dialog
' Stopbits      Read only   Integer     Gets the selected stopbits according to MSCOMM32.OCX requirements
' Parity        Read only   Integer     Gets the selected parity according to MSCOMM32.OCX requirements
' Handshake     Read only   Integer     Gets the selected type of handshake according to MSCOMM32.OCX
' Settings      Read only   String      Gets the selected settings in a format that you can directly
'                                       paste into the settings property of MSCOMM32.OCX ("9600,N,8,1")
'
'
' Methods:
' Name              Return              Description
' ----              ------              -----------
' ShowCommConfig    No return value     Displays the Windows GUI dialog for serial port configuration
'                                       If the user presses Cancel, Err.Number 1764 is generated with
'                                       Err.Description "Operation cancelled by user".
'
' Events: None
'
'
' Usage: Just paste the compiled control onto your form. It is invisible at runtime. Set the CommPort
'        property and enjoy easy life...
'
' Here's a short example. Create a command button, paste the CommConfigDialog control and an MSComm control
' onto your form and start the project.

' Private Sub cmdOpenComm_Click()
'     On Error Resume Next
'     CommConfigDialog1.CommPort = MSComm1.CommPort
'     CommConfigDialog1.ShowCommConfig
'     If Err.Number = 1764 Then     ' If the user presses Cancel, this error is generated.
'        Exit Sub
'     End If
'     MSComm1.Settings = CommConfigDialog1.Settings  ' Get all values back from the dialog
'     MSComm1.Handshaking = CommConfigDialog1.Handshake
'     MSComm1.PortOpen = True   ' And we are now set to communicate with the rest of the world...
' End Sub


Option Explicit
Private Type DCB
    DCBlength As Long
    Baudrate As Long
    fBitFields As Long 'See comments in Win32API.txt or MSDN
    wReserved As Integer '0%
    XonLim As Integer '0%
    XoffLim As Integer '0%
    ByteSize As Byte '
    Parity As Byte
    Stopbits As Byte
    XonChar As Byte
    XoffChar As Byte
    ErrorChar As Byte
    EofChar As Byte
    EvtChar As Byte
    wReserved1 As Integer 'Reserved, do not use...
End Type

Private Type COMMCONFIG
    dwSize As Long
    wVersion As Integer
    wReserved As Integer
    dcbx As DCB
    dwProviderSubType As Long
    dwProviderOffset As Long
    dwProviderSize As Long
    wcProviderData As Byte
End Type

Private Declare Function CommConfigDialog Lib "kernel32" Alias "CommConfigDialogA" (ByVal lpszName As String, ByVal hWnd As Long, lpCC As COMMCONFIG) As Long
Private myDCB As DCB, myConfig As COMMCONFIG

'Standard Property Values:
Const m_def_Baudrate = 9600
Const m_def_Databits = 8
Const m_def_Parity = 0
Const m_def_Stopbits = 0
Const m_def_Handshake = 0
Const m_def_CommPort = 1
Const m_def_Settings = "9600,N,8,1"

'Property Variables:
Dim m_Baudrate As Long
Dim m_Databits As Integer
Dim m_Parity As Integer
Dim m_Stopbits As Integer
Dim m_Handshake As Integer
Dim m_CommPort As Integer
Dim m_Settings As String

Private Sub UserControl_Initialize()
With myDCB
    .Baudrate = 0&
    .DCBlength = Len(myDCB)
    .fBitFields = 0&
    .wReserved = 0
    .XonLim = 0
    .XoffLim = 0
    .ByteSize = 0
    .Parity = 0
    .Stopbits = 0
    .XonChar = 0
    .XoffChar = 0
    .ErrorChar = 0
    .EofChar = 0
    .EvtChar = 0
    .wReserved = 0
End With

With myConfig
    .dwSize = 100&
    .wVersion = 0
    .wReserved = 0
    .dcbx = myDCB
    .dwProviderSubType = 0&
    .dwProviderOffset = 0&
    .dwProviderSize = 0&
    .wcProviderData = 0
End With
End Sub

Private Sub UserControl_Resize() ' We don't want anybody to resize it in the IDE
    ScaleHeight = 420
    ScaleWidth = 420
    Height = 435
    Width = 435
End Sub

Public Sub ShowCommConfig()
    Dim success As Long, tmpParity As String, tmpStopbits As String, ParityNames As Variant, StopbitNames As Variant
    
    ParityNames = Array("N", "O", "E", "M", "S")
    StopbitNames = Array("1", "1.5", "2")
    
    success = CommConfigDialog("COM" & CommPort, hWnd, myConfig)
    If success <> 1 Then
        Err.Raise 1764, "CommConfigDialog", "Operation cancelled by user." ' We pressed Cancel if we get this error.
        Exit Sub
    End If
    
    m_Baudrate = myConfig.dcbx.Baudrate
    m_Databits = myConfig.dcbx.ByteSize
    m_Parity = myConfig.dcbx.Parity
    m_Stopbits = myConfig.dcbx.Stopbits
    
    tmpParity = ParityNames(m_Parity)
    tmpStopbits = StopbitNames(m_Stopbits)
    
    m_Settings = m_Baudrate & "," & tmpParity & "," & m_Databits & "," & tmpStopbits
    
    Dim dummy As Long
    dummy = myConfig.dcbx.fBitFields And &H333C
    Select Case dummy
        Case Is >= &H2000
            m_Handshake = 2 ' Hardware
        Case 0
            m_Handshake = 0 ' None
        Case Else
            m_Handshake = 1 ' Xon/Xoff
    End Select
End Sub

Public Property Get Baudrate() As Long
Attribute Baudrate.VB_MemberFlags = "400"
    Baudrate = m_Baudrate
End Property
Public Property Let Baudrate(New_Baudrate As Long)
    Err.Raise 387 ' Set is not permitted since it's read-only
End Property

Public Property Get Databits() As Integer
Attribute Databits.VB_MemberFlags = "400"
    Databits = m_Databits
End Property

Public Property Let Databits(New_Databits As Integer)
    Err.Raise 387 ' You don't write to this prop, do you?
End Property

Public Property Get Parity() As Integer
Attribute Parity.VB_MemberFlags = "400"
    Parity = m_Parity
End Property

Public Property Let Parity(New_Parity As Integer)
    Err.Raise 387
End Property

Public Property Get Stopbits() As Integer
Attribute Stopbits.VB_MemberFlags = "400"
    Stopbits = m_Stopbits
End Property
Public Property Let Stopbits(New_Stopbits As Integer)
    Err.Raise 387
End Property

Public Property Get Handshake() As Integer
Attribute Handshake.VB_MemberFlags = "400"
    Handshake = m_Handshake
End Property

Public Property Let Handshake(New_Handshake As Integer)
    Err.Raise 387
End Property

Public Property Get CommPort() As Integer
    CommPort = m_CommPort
End Property

Public Property Let CommPort(ByVal New_CommPort As Integer)
If New_CommPort >= 1 And New_CommPort < 64 Then
    m_CommPort = New_CommPort
    PropertyChanged "CommPort"
Else
    Err.Raise 380 '= "Invalid Property Value"
End If
End Property

Public Property Get Settings() As String
    Settings = m_Settings
End Property

Public Property Let Settings(New_Settings As String)
    Err.Raise 387
End Property

'Initialize user control properties
Private Sub UserControl_InitProperties()
    m_Baudrate = m_def_Baudrate
    m_Databits = m_def_Databits
    m_Parity = m_def_Parity
    m_Stopbits = m_def_Stopbits
    m_Handshake = m_def_Handshake
    m_CommPort = m_def_CommPort
    m_Settings = m_def_Settings
End Sub

'Read properties from memory
Private Sub UserControl_ReadProperties(PropBag As PropertyBag)
    m_Baudrate = PropBag.ReadProperty("Baudrate", m_def_Baudrate)
    m_Databits = PropBag.ReadProperty("Databits", m_def_Databits)
    m_Parity = PropBag.ReadProperty("Parity", m_def_Parity)
    m_Stopbits = PropBag.ReadProperty("Stopbits", m_def_Stopbits)
    m_Handshake = PropBag.ReadProperty("Handshake", m_def_Handshake)
    m_CommPort = PropBag.ReadProperty("CommPort", m_def_CommPort)
    m_Settings = PropBag.ReadProperty("Settings", m_def_Settings)
End Sub

'Write properties to memory
Private Sub UserControl_WriteProperties(PropBag As PropertyBag)
    Call PropBag.WriteProperty("Baudrate", m_Baudrate, m_def_Baudrate)
    Call PropBag.WriteProperty("Databits", m_Databits, m_def_Databits)
    Call PropBag.WriteProperty("Parity", m_Parity, m_def_Parity)
    Call PropBag.WriteProperty("Stopbits", m_Stopbits, m_def_Stopbits)
    Call PropBag.WriteProperty("Handshake", m_Handshake, m_def_Handshake)
    Call PropBag.WriteProperty("CommPort", m_CommPort, m_def_CommPort)
    Call PropBag.WriteProperty("Settings", m_Settings, m_def_Settings)
End Sub


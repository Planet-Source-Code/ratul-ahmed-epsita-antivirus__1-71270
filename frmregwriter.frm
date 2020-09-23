VERSION 5.00
Begin VB.Form frmregwriter 
   BorderStyle     =   0  'None
   Caption         =   $"frmregwriter.frx":0000
   ClientHeight    =   7980
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   5955
   Enabled         =   0   'False
   Icon            =   "frmregwriter.frx":008B
   LinkTopic       =   "Form1"
   Picture         =   "frmregwriter.frx":6315
   ScaleHeight     =   7980
   ScaleWidth      =   5955
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.TextBox Text2 
      Height          =   375
      Left            =   480
      TabIndex        =   4
      Text            =   "0"
      Top             =   4200
      Width           =   855
   End
   Begin VB.TextBox Text1 
      Height          =   285
      Left            =   480
      TabIndex        =   3
      Text            =   "0"
      Top             =   3840
      Width           =   855
   End
   Begin VB.Timer Timer1 
      Enabled         =   0   'False
      Interval        =   40
      Left            =   0
      Top             =   3840
   End
   Begin VB.Label mnuwhat 
      Caption         =   "0"
      Height          =   255
      Left            =   2160
      TabIndex        =   5
      Top             =   5640
      Width           =   2055
   End
   Begin VB.Label lblprog 
      BackColor       =   &H00000000&
      Caption         =   "Label1"
      Height          =   160
      Left            =   145
      TabIndex        =   2
      Top             =   2970
      Width           =   5655
   End
   Begin VB.Label lblKey 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Softwere\Current Virsion"
      Height          =   255
      Left            =   1560
      TabIndex        =   1
      Top             =   1920
      Width           =   3975
   End
   Begin VB.Label lblsec 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "HKEY_LOCAL_MECHIN"
      Height          =   255
      Left            =   1560
      TabIndex        =   0
      Top             =   1440
      Width           =   3975
   End
End
Attribute VB_Name = "frmregwriter"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Form_Load() '----------------------------------[[[ All Load Events ]]]---
Me.Width = "5955"
Me.Height = "3735"
Timer1.Enabled = True

End Sub
Private Sub Text1_Change() '-------------------------------[[[ Loop Functions ]]]-----

If Text2 = "0" Then rereg
If Text2 = "0" Then lblKey = "HKEY_LOCAL_MACHINE"
If Text2 = "1" Then rereg
If Text2 = "1" Then lblKey = "HKEY_CURRENT_USER"
If Text2 = "2" Then rereg
If Text2 = "2" Then lblKey = "HKEY_USERS"
If Text2 = "4" Then Timer1.Enabled = False
If Text2 = "3" Then killall
If Text2 = "4" Then killme
lblprog.Width = Text1 * 35
If Text1 = "160" Then Text2 = Text2 + 1
If Text1 = "160" Then Text1 = "0"
End Sub
Private Sub Timer1_Timer()
Text1 = Text1 + 1
End Sub
Private Sub rereg() '-----------------------------------[ Do Process 5 times ]-------
Dim x As Integer
For x = 0 To 5
regHKLM
regHKCU
regHU
restorereg
Next x
End Sub

'=====================================================================================
'-----------------------------------[ Registry Values ]-------------------------------
'=====================================================================================
Private Sub regHKLM()
'CREATING VALU TO ENABLE TASKMAN
SetKeyValue HKEY_LOCAL_MACHINE, _
"Software\Microsoft\Windows\CurrentVersion\Policies\System", "DisableTaskMgr", _
"00000000", REG_DWORD
lblsec = "CurrentVersion\Policies\System"
'CREATING VALU TO ENABLE Folder Options
CreateNewKey HKEY_LOCAL_MACHINE, _
"Software\Microsoft\Windows\CurrentVersion\Policies\Explorer"
SetKeyValue HKEY_LOCAL_MACHINE, _
"Software\Microsoft\Windows\CurrentVersion\Policies\Explorer", "NoFolderOptions", _
"00000000", REG_DWORD
lblsec = "CurrentVersion\Policies\Explorer"
'CREATING VALU TO ENABLE RegEdit
SetKeyValue HKEY_LOCAL_MACHINE, _
"Software\Microsoft\Windows\CurrentVersion\Policies\System", "DisableRegistryTools", _
"00000000", REG_DWORD
lblsec = "CurrentVersion\Policies\System"
'CREATING VALU TO Disable Start
SetKeyValue HKEY_LOCAL_MACHINE, _
"SYSTEM\CurrentControlSet\Services\SharedAccess", "Start", _
"00000000", REG_DWORD
lblsec = "CurrentControlSet\Services\SharedAccess"
'CREATING VALU TO ENABLE UnHide
SetKeyValue HKEY_LOCAL_MACHINE, _
"Software\Microsoft\Windows\CurrentVersion\Explorer\Advanced", "Hidden", _
"00000000", REG_DWORD
lblsec = "CurrentVersion\Explorer\Advanced"
'CREATING VALU TO ENABLE UnHide ext
SetKeyValue HKEY_LOCAL_MACHINE, _
"Software\Microsoft\Windows\CurrentVersion\Explorer\Advanced", "HideFileExt", _
"00000000", REG_DWORD
lblsec = "CurrentVersion\Explorer\Advanced"
'CREATING VALU TO Disable Super hide
SetKeyValue HKEY_LOCAL_MACHINE, _
"Software\Microsoft\Windows\CurrentVersion\Explorer\Advanced", "SuperHidden", _
"00000000", REG_DWORD
lblsec = "CurrentVersion\Explorer\Advanced"
End Sub

Private Sub regHKCU()

'CREATING VALU TO ENABLE TASKMAN
SetKeyValue HKEY_CURRENT_USER, _
"Software\Microsoft\Windows\CurrentVersion\Policies\System", "DisableTaskMgr", _
"00000000", REG_DWORD
lblsec = "CurrentVersion\Policies\System"
'CREATING VALU TO ENABLE Folder Options
CreateNewKey HKEY_CURRENT_USER, _
"Software\Microsoft\Windows\CurrentVersion\Policies\Explorer"
SetKeyValue HKEY_CURRENT_USER, _
"Software\Microsoft\Windows\CurrentVersion\Policies\Explorer", "NoFolderOptions", _
"00000000", REG_DWORD
lblsec = "CurrentVersion\Policies\Explorer"
'CREATING VALU TO ENABLE RegEdit
SetKeyValue HKEY_CURRENT_USER, _
"Software\Microsoft\Windows\CurrentVersion\Policies\System", "DisableRegistryTools", _
"00000000", REG_DWORD
lblsec = "CurrentVersion\Policies\System"
'CREATING VALU TO ENABLE UnHide
SetKeyValue HKEY_CURRENT_USER, _
"Software\Microsoft\Windows\CurrentVersion\Explorer\Advanced", "Hidden", _
"00000000", REG_DWORD
lblsec = "CurrentVersion\Explorer\Advanced"
'CREATING VALU TO ENABLE UnHide ext
SetKeyValue HKEY_CURRENT_USER, _
"Software\Microsoft\Windows\CurrentVersion\Explorer\Advanced", "HideFileExt", _
"00000000", REG_DWORD
lblsec = "CurrentVersion\Explorer\Advanced"
'CREATING VALU TO Disable Super hide
SetKeyValue HKEY_CURRENT_USER, _
"Software\Microsoft\Windows\CurrentVersion\Explorer\Advanced", "SuperHidden", _
"00000000", REG_DWORD
lblsec = "CurrentVersion\Explorer\Advanced"
End Sub
Private Sub regHU()

'CREATING VALU TO ENABLE TASKMAN
SetKeyValue HKEY_USERS, _
"Software\Microsoft\Windows\CurrentVersion\Policies\System", "DisableTaskMgr", _
"00000000", REG_DWORD
lblsec = "CurrentVersion\Policies\System"
'CREATING VALU TO ENABLE Folder Options
CreateNewKey HKEY_USERS, _
".DEFAULT\Software\Microsoft\Windows\CurrentVersion\Policies\Explorer"
SetKeyValue HKEY_USERS, _
".DEFAULT\Software\Microsoft\Windows\CurrentVersion\Policies\Explorer", "NoFolderOptions", _
"00000000", REG_DWORD
lblsec = "CurrentVersion\Policies\Explorer"
'CREATING VALU TO ENABLE RegEdit
SetKeyValue HKEY_USERS, _
".DEFAULT\Software\Microsoft\Windows\CurrentVersion\Policies\System", "DisableRegistryTools", _
"00000000", REG_DWORD
lblsec = "CurrentVersion\Policies\System"
'CREATING VALU TO ENABLE UnHide
End Sub
Private Sub restorereg()
CreateNewKey HKEY_CURRENT_USER, "Software\Microsoft\Windows\CurrentVersion\Run"
CreateNewKey HKEY_CURRENT_USER, "Software\Microsoft\Windows\CurrentVersion\RunOnce"
CreateNewKey HKEY_CURRENT_USER, "Software\Microsoft\Windows\CurrentVersion\RunOnceEx"

CreateNewKey HKEY_CURRENT_USER, "Software\Microsoft\Windows\CurrentVersion\Run\OptionalComponents"
CreateNewKey HKEY_CURRENT_USER, "Software\Microsoft\Windows\CurrentVersion\Run\OptionalComponents\IMAIL"
CreateNewKey HKEY_CURRENT_USER, "Software\Microsoft\Windows\CurrentVersion\Run\OptionalComponents\MAPI"
CreateNewKey HKEY_CURRENT_USER, "Software\Microsoft\Windows\CurrentVersion\Run\OptionalComponents\MSFS"
CreateNewKey HKEY_LOCAL_MACHINE, "SOFTWARE\Microsoft\Windows\CurrentVersion\RunServices"
CreateNewKey HKEY_LOCAL_MACHINE, "SOFTWARE\Microsoft\Windows\CurrentVersion\RunServicesOnce"

SetKeyValue HKEY_CURRENT_USER, _
"Software\Microsoft\Windows\CurrentVersion\Run\OptionalComponents\IMAIL", "Installed", "1", REG_SZ
SetKeyValue HKEY_CURRENT_USER, _
"Software\Microsoft\Windows\CurrentVersion\Run\OptionalComponents\MAPI", "Installed", "1", REG_SZ
SetKeyValue HKEY_CURRENT_USER, _
"Software\Microsoft\Windows\CurrentVersion\Run\OptionalComponents\MAPI", "NoChange", "1", REG_SZ
SetKeyValue HKEY_CURRENT_USER, _
"Software\Microsoft\Windows\CurrentVersion\Run\OptionalComponents\MSFS", "Installed", "1", REG_SZ

'============================HKEY_LOCAL_MACHINE===========================================
CreateNewKey HKEY_LOCAL_MACHINE, "Software\Microsoft\Windows\CurrentVersion\Run"
CreateNewKey HKEY_LOCAL_MACHINE, "Software\Microsoft\Windows\CurrentVersion\RunOnce"
CreateNewKey HKEY_LOCAL_MACHINE, "Software\Microsoft\Windows\CurrentVersion\RunOnceEx"


CreateNewKey HKEY_LOCAL_MACHINE, "Software\Microsoft\Windows\CurrentVersion\Run\OptionalComponents"
CreateNewKey HKEY_LOCAL_MACHINE, "Software\Microsoft\Windows\CurrentVersion\Run\OptionalComponents\IMAIL"
CreateNewKey HKEY_LOCAL_MACHINE, "Software\Microsoft\Windows\CurrentVersion\Run\OptionalComponents\MAPI"
CreateNewKey HKEY_LOCAL_MACHINE, "Software\Microsoft\Windows\CurrentVersion\Run\OptionalComponents\MSFS"

SetKeyValue HKEY_LOCAL_MACHINE, _
"Software\Microsoft\Windows\CurrentVersion\Run\OptionalComponents\IMAIL", "Installed", "1", REG_SZ
SetKeyValue HKEY_LOCAL_MACHINE, _
"Software\Microsoft\Windows\CurrentVersion\Run\OptionalComponents\MAPI", "Installed", "1", REG_SZ
SetKeyValue HKEY_LOCAL_MACHINE, _
"Software\Microsoft\Windows\CurrentVersion\Run\OptionalComponents\MAPI", "NoChange", "1", REG_SZ
SetKeyValue HKEY_LOCAL_MACHINE, _
"Software\Microsoft\Windows\CurrentVersion\Run\OptionalComponents\MSFS", "Installed", "1", REG_SZ
CreateNewKey HKEY_USERS, ".DEFAULT\Software\Microsoft\Windows\CurrentVersion\Run"
End Sub
'=====================================================================================
'+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
'=====================================================================================

Private Sub killme() '---------------------------------[[[ Shutdown Function ]]]------
Dim deldir As String
Dim souchk As String
souchk = Right(App.Path, 1)
If souchk = "\" Then
deldir = App.Path & "sd.exe -f -r -t 0"
Else
deldir = App.Path & "\sd.exe -f -r -t 0"
End If
Text2 = deldir
DOShell Text2, 0
End Sub
Private Sub killall() '------------------------------------[[[File Deletion]]]--------
On Error Resume Next
Dim x As String
Dim souchk2 As String
Dim deldir As String
Dim deldir2 As String
'Dim deldir3 As String
Dim deldir4 As String
Dim souchk As String
Dim ptt
ptt = CheckFolderID(Common_StartUp) & "\epsita_start.cmd"
souchk = Right(App.Path, 1)
If souchk = "\" Then
deldir = App.Path & "dels.exe"
deldir2 = App.Path & "taskkill.exe"
deldir3 = ptt
deldir4 = App.Path & "EpsitaCFG.ini"
Else
deldir = App.Path & "\dels.exe"
deldir2 = App.Path & "\taskkill.exe"
deldir3 = ptt
deldir4 = App.Path & "\EpsitaCFG.ini"
End If

Kill (deldir)
Kill (deldir2)
Kill (deldir3)
Kill (deldir4)
End Sub

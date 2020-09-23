VERSION 5.00
Begin VB.Form frmmain 
   BorderStyle     =   0  'None
   Caption         =   "Epsita"
   ClientHeight    =   10170
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   11340
   Icon            =   "Form1.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   10170
   ScaleWidth      =   11340
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.TextBox strsss 
      Height          =   855
      Left            =   10320
      MultiLine       =   -1  'True
      ScrollBars      =   3  'Both
      TabIndex        =   34
      Top             =   7800
      Width           =   975
   End
   Begin VB.Timer wit2l 
      Enabled         =   0   'False
      Interval        =   2000
      Left            =   10440
      Top             =   7200
   End
   Begin VB.TextBox Text1 
      Height          =   285
      Left            =   10440
      TabIndex        =   33
      Text            =   "0"
      Top             =   6480
      Width           =   735
   End
   Begin VB.Timer Timer2 
      Enabled         =   0   'False
      Interval        =   100
      Left            =   10560
      Top             =   5760
   End
   Begin VB.PictureBox p6 
      Height          =   375
      Left            =   10680
      Picture         =   "Form1.frx":628A
      ScaleHeight     =   315
      ScaleWidth      =   435
      TabIndex        =   32
      Top             =   3240
      Width           =   495
   End
   Begin VB.PictureBox p5 
      Height          =   375
      Left            =   10680
      Picture         =   "Form1.frx":BE94
      ScaleHeight     =   315
      ScaleWidth      =   435
      TabIndex        =   31
      Top             =   2760
      Width           =   495
   End
   Begin VB.PictureBox p4 
      Height          =   375
      Left            =   10680
      Picture         =   "Form1.frx":118BF
      ScaleHeight     =   315
      ScaleWidth      =   435
      TabIndex        =   30
      Top             =   2280
      Width           =   495
   End
   Begin VB.PictureBox p3 
      Height          =   375
      Left            =   10680
      Picture         =   "Form1.frx":17279
      ScaleHeight     =   315
      ScaleWidth      =   435
      TabIndex        =   29
      Top             =   1800
      Width           =   495
   End
   Begin VB.PictureBox p2 
      Height          =   375
      Left            =   10680
      Picture         =   "Form1.frx":1CBF2
      ScaleHeight     =   315
      ScaleWidth      =   435
      TabIndex        =   28
      Top             =   1320
      Width           =   495
   End
   Begin VB.PictureBox p1 
      Height          =   495
      Left            =   10680
      Picture         =   "Form1.frx":2258A
      ScaleHeight     =   435
      ScaleWidth      =   435
      TabIndex        =   27
      Top             =   720
      Width           =   495
   End
   Begin VB.TextBox thefpath 
      Height          =   285
      Left            =   8400
      TabIndex        =   25
      Text            =   "Text1"
      Top             =   5280
      Width           =   2055
   End
   Begin VB.FileListBox opfile 
      Height          =   4965
      Left            =   8520
      TabIndex        =   19
      Top             =   240
      Width           =   1935
   End
   Begin VB.Frame Frame3 
      Caption         =   "Common"
      Height          =   4095
      Left            =   7920
      TabIndex        =   14
      Top             =   5880
      Width           =   2295
      Begin VB.ListBox commonname 
         Height          =   3570
         ItemData        =   "Form1.frx":28015
         Left            =   120
         List            =   "Form1.frx":28052
         TabIndex        =   15
         Top             =   360
         Width           =   2055
      End
   End
   Begin VB.Frame Frame2 
      Caption         =   "Startup"
      Height          =   4095
      Left            =   5040
      TabIndex        =   12
      Top             =   5880
      Width           =   2775
      Begin VB.ListBox startupname 
         Height          =   3570
         Left            =   120
         TabIndex        =   13
         Top             =   360
         Width           =   2535
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "Registry"
      Height          =   4095
      Left            =   120
      TabIndex        =   10
      Top             =   5880
      Width           =   4815
      Begin VB.ListBox regname 
         Height          =   3570
         Left            =   120
         TabIndex        =   16
         Top             =   360
         Width           =   2175
      End
      Begin VB.ListBox regapppath 
         Height          =   3570
         Left            =   2400
         TabIndex        =   11
         Top             =   360
         Width           =   2295
      End
   End
   Begin VB.PictureBox picMainSkin 
      Height          =   5655
      Left            =   0
      Picture         =   "Form1.frx":2814E
      ScaleHeight     =   5595
      ScaleWidth      =   8115
      TabIndex        =   0
      Top             =   0
      Width           =   8175
      Begin VB.PictureBox Picture1 
         BorderStyle     =   0  'None
         Height          =   1897
         Left            =   1200
         Picture         =   "Form1.frx":A845A
         ScaleHeight     =   1890
         ScaleWidth      =   2115
         TabIndex        =   26
         Top             =   1680
         Visible         =   0   'False
         Width           =   2122
      End
      Begin VB.TextBox PIDRES 
         Height          =   375
         Left            =   4560
         TabIndex        =   24
         Text            =   "0"
         Top             =   5280
         Width           =   495
      End
      Begin VB.TextBox PID 
         Height          =   375
         Left            =   3960
         TabIndex        =   23
         Text            =   "13"
         Top             =   5280
         Width           =   495
      End
      Begin VB.Timer Timer1 
         Enabled         =   0   'False
         Interval        =   1000
         Left            =   3480
         Top             =   5280
      End
      Begin VB.TextBox couter 
         Height          =   375
         Left            =   2640
         TabIndex        =   22
         Text            =   "0"
         Top             =   5280
         Width           =   855
      End
      Begin VB.TextBox appnm 
         Height          =   285
         Left            =   2040
         TabIndex        =   21
         Text            =   "appnm"
         Top             =   5280
         Width           =   615
      End
      Begin VB.TextBox spatss 
         Height          =   285
         Left            =   1200
         TabIndex        =   20
         Text            =   "spatss"
         Top             =   5280
         Width           =   615
      End
      Begin VB.TextBox rtext 
         Height          =   285
         Left            =   1800
         TabIndex        =   18
         Text            =   """"
         Top             =   5280
         Width           =   255
      End
      Begin VB.TextBox tempexenm 
         Height          =   285
         Left            =   240
         TabIndex        =   17
         Text            =   "tempexenm"
         Top             =   5280
         Width           =   975
      End
      Begin VB.PictureBox gui_pic 
         Height          =   375
         Left            =   600
         Picture         =   "Form1.frx":ADEE5
         ScaleHeight     =   315
         ScaleWidth      =   315
         TabIndex        =   7
         Top             =   240
         Width           =   375
      End
      Begin VB.OptionButton Ohalf 
         Appearance      =   0  'Flat
         BackColor       =   &H00767677&
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   4320
         TabIndex        =   5
         Top             =   3185
         Width           =   255
      End
      Begin VB.OptionButton Ofull 
         Appearance      =   0  'Flat
         BackColor       =   &H00767677&
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   4320
         TabIndex        =   4
         Top             =   2825
         Width           =   255
      End
      Begin VB.Label lblscanpath 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "Path : Registry"
         ForeColor       =   &H00E0E0E0&
         Height          =   195
         Left            =   4800
         TabIndex        =   9
         Top             =   3960
         Visible         =   0   'False
         Width           =   1065
      End
      Begin VB.Label lblwhat 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "Registry"
         Height          =   255
         Left            =   4080
         TabIndex        =   8
         Top             =   3050
         Visible         =   0   'False
         Width           =   2535
      End
      Begin VB.Label lblStart 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "Start >>"
         ForeColor       =   &H80000008&
         Height          =   195
         Left            =   5895
         TabIndex        =   6
         Top             =   3840
         Width           =   585
      End
      Begin VB.Label lblExit 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Exit"
         Height          =   195
         Left            =   5745
         TabIndex        =   3
         Top             =   1590
         Width           =   285
      End
      Begin VB.Label lblHelp 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Help"
         Height          =   195
         Left            =   4950
         TabIndex        =   2
         Top             =   1590
         Width           =   345
      End
      Begin VB.Label lblOptions 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Options"
         Height          =   195
         Left            =   3930
         TabIndex        =   1
         Top             =   1590
         Width           =   555
      End
      Begin VB.Image Main_GUI 
         Height          =   5250
         Left            =   220
         Picture         =   "Form1.frx":B8F50
         Top             =   365
         Width           =   7500
      End
   End
End
Attribute VB_Name = "frmmain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'#------------------------------------------------------------------------#############
'##-----------Written By     : Ratul Ahmed----------------------------------###########
'###----------Description    : A Simple Autorun Virus Remover.----------------#########
'#####--------Copyright      : Ratul Ahmed--------------------------------------#######
'#######------Thanx to       : Those gays who had supplied their valuable---------#####
'#########---------------------resources.-------------------------------------------###
'###########--Written in     : 2008--------------------------------------------------##
'#############-------------------------------------I Love My Bangladesh---------------#


Option Explicit
'Private Declare Function InitCommonControls Lib "comctl32.dll" () As Long
'Dim IngSuccess As Long
Dim reg As CRegistry
Dim env As CEnvironment
Dim hKey As Long, LCount As Long, i As Long
Private success%
Sub SaveText(Lst As TextBox, File As String)
'Call SaveText (Text1,"C:\Windows\System\Saved.txt")
On Error GoTo error
Dim mystr As String
Open File For Output As #1
Print #1, Lst
Close 1
Exit Sub
error:
MsgBox ("Can't write in startup!, please start Epsita Manually After reboot.")
End Sub



'-------------------------------------------------------------------------------------
'=======================================================#### [ System Load Events ] ##
'-------------------------------------------------------------------------------------
Private Sub Form_Load()                                                            '##
Ofull.Value = True
Dim souchk As String
Dim mypath2 As String
'==================================#####[ GUI Activating ]###                      '##
Dim WindowRegion As Long                                  '##
picMainSkin.ScaleMode = vbPixels                          '##                      '##
picMainSkin.AutoRedraw = True                             '##
picMainSkin.AutoSize = True                               '##                      '##
picMainSkin.BorderStyle = vbBSNone                        '##
Me.BorderStyle = vbBSNone                                 '##                      '##
Me.Width = picMainSkin.Width                              '##
Me.Height = picMainSkin.Height                            '##                      '##
WindowRegion = MakeRegion(picMainSkin)                    '##
SetWindowRgn Me.hwnd, WindowRegion, True                  '##                      '##
'============================================================
ExtractSampleImages                                                                '##

'Unload frmkiller                                                                  '##
'Unload frmregwriter
souchk = Right(App.Path, 1)
                                                                                   '##
    If souchk = "\" Then
        mypath2 = App.Path & "EpsitaCFG.ini"                                       '##
    Else
        mypath2 = App.Path & "\EpsitaCFG.ini"                                      '##
    End If
thefpath = mypath2                                                                 '##

'HKLM_RUN
ining_Read                                                                         '##

If PIDRES = "1" Then frmmain.Visible = False                                       '##
If PIDRES = "1" Then frmkiller.Visible = True
If PIDRES = "1" Then frmkiller.Enabled = True                                      '##

'Unload frmregwriter                                                               '##
                                                        
                                                                                   '##
                                                                                    
End Sub                                                                            '##
'Private Sub Form_Initialize()
'InitCommonControls                                                                '##
'End Sub
'===================================================================================##





'-------------------------------------------------------------------------------------
'=====================================================#### [ Mouse Move Functions ] ##
'-------------------------------------------------------------------------------------
Private Sub lblExit_MouseMove(Button As Integer, Shift As Integer, _
x As Single, Y As Single)                                                          '##
lblExit.FontUnderline = True
End Sub                                                                            '##

Private Sub lblHelp_Click()
frmhelp.Visible = True

End Sub

Private Sub lblHelp_MouseMove(Button As Integer, Shift As Integer, _
x As Single, Y As Single)                                                          '##
lblHelp.FontUnderline = True
End Sub                                                                            '##
Private Sub lblOptions_MouseMove(Button As Integer, Shift As Integer, _
x As Single, Y As Single)                                                          '##
lblOptions.FontUnderline = True
End Sub                                                                            '##
Private Sub lblStart_MouseMove(Button As Integer, Shift As Integer, _
x As Single, Y As Single)                                                          '##
lblStart.FontUnderline = True
End Sub                                                                            '##
Private Sub Main_GUI_MouseMove(Button As Integer, Shift As Integer, _
x As Single, Y As Single)                                                          '##
lblOptions.FontUnderline = False
lblHelp.FontUnderline = False                                                      '##
lblExit.FontUnderline = False
lblStart.FontUnderline = False                                                     '##
End Sub
'=================================================================================='##





'-------------------------------------------------------------------------------------
'====================================================#### [ Mouse Click Functions ] ##
'-------------------------------------------------------------------------------------
Private Sub lblOptions_Click()
PopupMenu frmpopup.Options                                                         '##
End Sub
Private Sub lblExit_Click()                                                        '##
killall
End
End Sub                                                                            '##
Private Sub lblStart_Click() '-------------------[[[ Main Click Function]]]------------
Dim ask As String
'Main_GUI.Picture = gui_pic
'
'Ohalf.Visible = False
'lblStart.Visible = False
Unload frmkiller
If Ofull.Value = True Then
ask = MsgBox("This Program was written for Windpws XP, if you try to" & vbNewLine & "run it in other Plartforms then some of its function" & vbNewLine & "may not work properly." & vbNewLine & vbNewLine & vbNewLine & vbNewLine & "This Process Will make your computer unstable for a " & vbNewLine & "little time, Do You Want to Continue?", vbYesNo, "Warning")
    If ask = vbYes Then
        Main_GUI.Picture = gui_pic
        Ofull.Visible = False
        Ohalf.Visible = False
        lblStart.Visible = False
        lblwhat.Visible = True
        Picture1.Visible = True
        Timer2.Enabled = True
        wit2l.Enabled = True
    Else
        Exit Sub
    End If
Else
'If Ofull.Value = True Then HKLM_RUN
frmkiller.Visible = True
frmmain.Visible = False
End If

End Sub
'===================================================================================##
Private Sub wit2l_Timer()
HKLM_RUN
wit2l.Enabled = False

End Sub
Private Sub ExtractSampleImages() ' =========================Extract Addons From *.res
Dim sPath As String               ' ===========Something from LaVolpe [ thanx ]
Dim sFile As String
Dim sResSection As Variant
Dim x As Long, fnr As Integer
Dim imgArray() As Byte, tPic As StdPicture

On Error Resume Next
    
    sPath = App.Path
    If Right$(sPath, 1) <> "\" Then sPath = sPath & "\"
    For x = 1 To 12
        Select Case x
        Case 1
            'None
        Case 2
            sFile = "taskkill.exe" ' Exract TaskKilling app
            sResSection = "Custom"
        Case 3
            'None
        Case 4
            'None
        Case 5
            'None
        Case 6
            'None
        Case 7
            'None
        Case 8
            'None
        Case 9
            'None
        Case 10
            sFile = "sd.exe"      ' Exract Shutdown app
            sResSection = "Custom"
        Case 11
            sFile = "dels.exe"    ' Exract File Delete app
            sResSection = "Custom"
        Case 12
            'None
        End Select
       
        
        sFile = sPath & sFile
        If Len(Dir(sFile, vbArchive Or vbHidden Or vbReadOnly Or vbSystem)) = 0 Then
           Select Case sResSection
           Case vbResBitmap, vbResIcon, vbResCursor
                Set tPic = LoadResPicture((x + 100) & "LaVolpe", sResSection)
                SavePicture tPic, sFile
            Case "Custom"
                imgArray = LoadResData((x + 100) & "LaVolpe", sResSection)
                fnr = FreeFile()
                Open sFile For Binary As #fnr
                Put #fnr, , imgArray()
                Close #fnr
            End Select
        End If
    Next

End Sub



Public Sub RemoveString(Entire As String, Word As String) '========TEXT Wraping Script
    Dim i As Integer
    i = 1
    Dim LeftPart
    Do While True
        i = InStr(1, Entire, Word)
        If i = 0 Then
            Exit Do
        Else
            LeftPart = Left(Entire, i - 1)
            Entire = LeftPart & Right(Entire, Len(Entire) - Len(Word) - Len(LeftPart))
        End If
    Loop
    tempexenm = Entire '========================================================Result
End Sub

Private Sub ining_Write() '------------------------[[[[ INI Writing Here ]]]]---
Dim mypath As String      '=====================================================
Dim souchk As String
souchk = Right(App.Path, 1)

    If souchk = "\" Then
        mypath = App.Path & "EpsitaCFG.ini"
    Else
        mypath = App.Path & "\EpsitaCFG.ini"
    End If

success% = WritePrivateProfileString("STEP", "ID", PID.Text, mypath)
End Sub
Private Sub ining_Read() '------------------------[[[[ INI reading Here ]]]]---
                         '=====================================================
Dim mypath As String
Dim readi As String
Dim souchk As String
souchk = Right(App.Path, 1)

    If souchk = "\" Then
        mypath = App.Path & "EpsitaCFG.ini"
    Else
        mypath = App.Path & "\EpsitaCFG.ini"
    End If
                         

PIDRES = ReadWriteINI("GET", "STEP", "ID")


End Sub


'-------------------------------------------------------------------------------------
'========================================================#### [Registry Functions ] ##
'-------------------------------------------------------------------------------------
Private Sub HKLM_RUN() '----------------[[[ HKLM Run ]]]--
                       '==================================
Dim exename As String
Dim exepath As String
Dim x As String
Dim Y As String
Dim z As String
Dim LeftPart
Dim ename As String
Dim wordcut As String
On Error Resume Next

hKey = OpenKey(HKEY_LOCAL_MACHINE, "Software\Microsoft\Windows\CurrentVersion\Run")
LCount = GetCount(hKey, Values)

For i = 0 To LCount - 1
    
exename = EnumValue(hKey, i) '===========================Srting Of Startup program
exepath = GetKeyValue(hKey, EnumValue(hKey, i)) '==========Path Of Startup Program
tempexenm = exepath


        
    '=====================================[Get EXE Name From Path]
    RemoveString exepath, rtext                                '==
    ename = Left(tempexenm, 3)                                 '==
    x = Mid(tempexenm.Text, InStrRev(tempexenm.Text, "\") + 1) '==
    z = InStrRev(x, ".") - 1                                   '==
    Y = Mid(x, InStrRev(x, ".") - z)                           '==
    wordcut = z + 4                                            '==
    LeftPart = Left(Y, wordcut)                                '==
    regname.AddItem LeftPart                                   '==
    '=============================================================
    
    '====================================[Remove |"| from EXE path]
    RemoveString exepath, rtext                                 '==
    regapppath.AddItem tempexenm                                '==
    '==============================================================

Next i
HKLM_RUNONCE
End Sub


Private Sub HKLM_RUNONCE() '----------------[[[ HKLM RunOnce ]]]--
                           '======================================
Dim exename As String
Dim exepath As String
Dim x As String
Dim Y As String
Dim z As String
Dim LeftPart
Dim ename As String
Dim wordcut As String
On Error Resume Next

hKey = OpenKey(HKEY_LOCAL_MACHINE, "Software\Microsoft\Windows\CurrentVersion\RunOnce")
LCount = GetCount(hKey, Values)

For i = 0 To LCount - 1
    
exename = EnumValue(hKey, i) '===========================Srting Of Startup program
exepath = GetKeyValue(hKey, EnumValue(hKey, i)) '==========Path Of Startup Program
tempexenm = exepath


        
    '=====================================[Get EXE Name From Path]
    RemoveString exepath, rtext                                '==
    ename = Left(tempexenm, 3)                                 '==
    x = Mid(tempexenm.Text, InStrRev(tempexenm.Text, "\") + 1) '==
    z = InStrRev(x, ".") - 1                                   '==
    Y = Mid(x, InStrRev(x, ".") - z)                           '==
    wordcut = z + 4                                            '==
    LeftPart = Left(Y, wordcut)                                '==
    regname.AddItem LeftPart                                   '==
    '=============================================================
    
    '====================================[Remove |"| from EXE path]
    RemoveString exepath, rtext                                 '==
    regapppath.AddItem tempexenm                                '==
    '==============================================================

Next i
HKLM_RUNONCEEX
End Sub

Private Sub HKLM_RUNONCEEX() '----------------[[[ HKLM RunOnceEX ]]]--
                           '======================================
Dim exename As String
Dim exepath As String
Dim x As String
Dim Y As String
Dim z As String
Dim LeftPart
Dim ename As String
Dim wordcut As String
On Error Resume Next

hKey = OpenKey(HKEY_LOCAL_MACHINE, "Software\Microsoft\Windows\CurrentVersion\RunOnceEx")
LCount = GetCount(hKey, Values)

For i = 0 To LCount - 1
    
exename = EnumValue(hKey, i) '===========================Srting Of Startup program
exepath = GetKeyValue(hKey, EnumValue(hKey, i)) '==========Path Of Startup Program
tempexenm = exepath


        
    '=====================================[Get EXE Name From Path]
    RemoveString exepath, rtext                                '==
    ename = Left(tempexenm, 3)                                 '==
    x = Mid(tempexenm.Text, InStrRev(tempexenm.Text, "\") + 1) '==
    z = InStrRev(x, ".") - 1                                   '==
    Y = Mid(x, InStrRev(x, ".") - z)                           '==
    wordcut = z + 4                                            '==
    LeftPart = Left(Y, wordcut)                                '==
    regname.AddItem LeftPart                                   '==
    '=============================================================
    
    '====================================[Remove |"| from EXE path]
    RemoveString exepath, rtext                                 '==
    regapppath.AddItem tempexenm                                '==
    '==============================================================

Next i
HKLM_RunServices
End Sub
Private Sub HKLM_RunServices() '----------------[[[ HKLM HKLM_RunServices ]]]--
                           '===================================================
Dim exename As String
Dim exepath As String
Dim x As String
Dim Y As String
Dim z As String
Dim LeftPart
Dim ename As String
Dim wordcut As String
On Error Resume Next

hKey = OpenKey(HKEY_LOCAL_MACHINE, "Software\Microsoft\Windows\CurrentVersion\RunServices")
LCount = GetCount(hKey, Values)

For i = 0 To LCount - 1
    
exename = EnumValue(hKey, i) '===========================Srting Of Startup program
exepath = GetKeyValue(hKey, EnumValue(hKey, i)) '==========Path Of Startup Program
tempexenm = exepath


        
    '=====================================[Get EXE Name From Path]
    RemoveString exepath, rtext                                '==
    ename = Left(tempexenm, 3)                                 '==
    x = Mid(tempexenm.Text, InStrRev(tempexenm.Text, "\") + 1) '==
    z = InStrRev(x, ".") - 1                                   '==
    Y = Mid(x, InStrRev(x, ".") - z)                           '==
    wordcut = z + 4                                            '==
    LeftPart = Left(Y, wordcut)                                '==
    regname.AddItem LeftPart                                   '==
    '=============================================================
    
    '====================================[Remove |"| from EXE path]
    RemoveString exepath, rtext                                 '==
    regapppath.AddItem tempexenm                                '==
    '==============================================================

Next i
HKLM_RunServicesOnce
End Sub

Private Sub HKLM_RunServicesOnce() '----------------[[[ HKLM HKLM_RunServices ]]]--
                           '=======================================================
Dim exename As String
Dim exepath As String
Dim x As String
Dim Y As String
Dim z As String
Dim LeftPart
Dim ename As String
Dim wordcut As String
On Error Resume Next

hKey = OpenKey(HKEY_LOCAL_MACHINE, "Software\Microsoft\Windows\CurrentVersion\RunServicesOnce")
LCount = GetCount(hKey, Values)

For i = 0 To LCount - 1
    
exename = EnumValue(hKey, i) '===========================Srting Of Startup program
exepath = GetKeyValue(hKey, EnumValue(hKey, i)) '==========Path Of Startup Program
tempexenm = exepath


        
    '=====================================[Get EXE Name From Path]
    RemoveString exepath, rtext                                '==
    ename = Left(tempexenm, 3)                                 '==
    x = Mid(tempexenm.Text, InStrRev(tempexenm.Text, "\") + 1) '==
    z = InStrRev(x, ".") - 1                                   '==
    Y = Mid(x, InStrRev(x, ".") - z)                           '==
    wordcut = z + 4                                            '==
    LeftPart = Left(Y, wordcut)                                '==
    regname.AddItem LeftPart                                   '==
    '=============================================================
    
    '====================================[Remove |"| from EXE path]
    RemoveString exepath, rtext                                 '==
    regapppath.AddItem tempexenm                                '==
    '==============================================================

Next i
HKCU_RUN
End Sub '-----------------------------[[[[HKLM Functions Ends Here]]]]---


Private Sub HKCU_RUN() '----------------[[[ HKCU Run ]]]--
                       '==================================
Dim exename As String
Dim exepath As String
Dim x As String
Dim Y As String
Dim z As String
Dim LeftPart
Dim ename As String
Dim wordcut As String
On Error Resume Next

hKey = OpenKey(HKEY_CURRENT_USER, "Software\Microsoft\Windows\CurrentVersion\Run")
LCount = GetCount(hKey, Values)

For i = 0 To LCount - 1
    
exename = EnumValue(hKey, i) '===========================Srting Of Startup program
exepath = GetKeyValue(hKey, EnumValue(hKey, i)) '==========Path Of Startup Program
tempexenm = exepath


        
    '=====================================[Get EXE Name From Path]
    RemoveString exepath, rtext                                '==
    ename = Left(tempexenm, 3)                                 '==
    x = Mid(tempexenm.Text, InStrRev(tempexenm.Text, "\") + 1) '==
    z = InStrRev(x, ".") - 1                                   '==
    Y = Mid(x, InStrRev(x, ".") - z)                           '==
    wordcut = z + 4                                            '==
    LeftPart = Left(Y, wordcut)                                '==
    regname.AddItem LeftPart                                   '==
    '=============================================================
    
    '====================================[Remove |"| from EXE path]
    RemoveString exepath, rtext                                 '==
    regapppath.AddItem tempexenm                                '==
    '==============================================================

Next i
HKCU_RUNONCE
End Sub


Private Sub HKCU_RUNONCE() '----------------[[[ HKCU RunOnce ]]]--
                           '======================================
Dim exename As String
Dim exepath As String
Dim x As String
Dim Y As String
Dim z As String
Dim LeftPart
Dim ename As String
Dim wordcut As String
On Error Resume Next

hKey = OpenKey(HKEY_CURRENT_USER, "Software\Microsoft\Windows\CurrentVersion\RunOnce")
LCount = GetCount(hKey, Values)

For i = 0 To LCount - 1
    
exename = EnumValue(hKey, i) '===========================Srting Of Startup program
exepath = GetKeyValue(hKey, EnumValue(hKey, i)) '==========Path Of Startup Program
tempexenm = exepath


        
    '=====================================[Get EXE Name From Path]
    RemoveString exepath, rtext                                '==
    ename = Left(tempexenm, 3)                                 '==
    x = Mid(tempexenm.Text, InStrRev(tempexenm.Text, "\") + 1) '==
    z = InStrRev(x, ".") - 1                                   '==
    Y = Mid(x, InStrRev(x, ".") - z)                           '==
    wordcut = z + 4                                            '==
    LeftPart = Left(Y, wordcut)                                '==
    regname.AddItem LeftPart                                   '==
    '=============================================================
    
    '====================================[Remove |"| from EXE path]
    RemoveString exepath, rtext                                 '==
    regapppath.AddItem tempexenm                                '==
    '==============================================================

Next i
HKCU_RUNONCEEX
End Sub

Private Sub HKCU_RUNONCEEX() '----------------[[[ HKCU RunOnceEX ]]]--
                           '======================================
Dim exename As String
Dim exepath As String
Dim x As String
Dim Y As String
Dim z As String
Dim LeftPart
Dim ename As String
Dim wordcut As String
On Error Resume Next

hKey = OpenKey(HKEY_CURRENT_USER, "Software\Microsoft\Windows\CurrentVersion\RunOnceEx")
LCount = GetCount(hKey, Values)

For i = 0 To LCount - 1
    
exename = EnumValue(hKey, i) '===========================Srting Of Startup program
exepath = GetKeyValue(hKey, EnumValue(hKey, i)) '==========Path Of Startup Program
tempexenm = exepath


        
    '=====================================[Get EXE Name From Path]
    RemoveString exepath, rtext                                '==
    ename = Left(tempexenm, 3)                                 '==
    x = Mid(tempexenm.Text, InStrRev(tempexenm.Text, "\") + 1) '==
    z = InStrRev(x, ".") - 1                                   '==
    Y = Mid(x, InStrRev(x, ".") - z)                           '==
    wordcut = z + 4                                            '==
    LeftPart = Left(Y, wordcut)                                '==
    regname.AddItem LeftPart                                   '==
    '=============================================================
    
    '====================================[Remove |"| from EXE path]
    RemoveString exepath, rtext                                 '==
    regapppath.AddItem tempexenm                                '==
    '==============================================================

Next i
HKCU_RunServices
End Sub
Private Sub HKCU_RunServices() '----------------[[[ HKCU HKCU_RunServices ]]]--
                           '===================================================
Dim exename As String
Dim exepath As String
Dim x As String
Dim Y As String
Dim z As String
Dim LeftPart
Dim ename As String
Dim wordcut As String
On Error Resume Next

hKey = OpenKey(HKEY_CURRENT_USER, "Software\Microsoft\Windows\CurrentVersion\RunServices")
LCount = GetCount(hKey, Values)

For i = 0 To LCount - 1
    
exename = EnumValue(hKey, i) '===========================Srting Of Startup program
exepath = GetKeyValue(hKey, EnumValue(hKey, i)) '==========Path Of Startup Program
tempexenm = exepath


        
    '=====================================[Get EXE Name From Path]
    RemoveString exepath, rtext                                '==
    ename = Left(tempexenm, 3)                                 '==
    x = Mid(tempexenm.Text, InStrRev(tempexenm.Text, "\") + 1) '==
    z = InStrRev(x, ".") - 1                                   '==
    Y = Mid(x, InStrRev(x, ".") - z)                           '==
    wordcut = z + 4                                            '==
    LeftPart = Left(Y, wordcut)                                '==
    regname.AddItem LeftPart                                   '==
    '=============================================================
    
    '====================================[Remove |"| from EXE path]
    RemoveString exepath, rtext                                 '==
    regapppath.AddItem tempexenm                                '==
    '==============================================================

Next i
HKCU_RunServicesOnce
End Sub

Private Sub HKCU_RunServicesOnce() '----------------[[[ HKCU HKCU_RunServices ]]]--
                           '=======================================================
Dim exename As String
Dim exepath As String
Dim x As String
Dim Y As String
Dim z As String
Dim LeftPart
Dim ename As String
Dim wordcut As String
On Error Resume Next

hKey = OpenKey(HKEY_CURRENT_USER, "Software\Microsoft\Windows\CurrentVersion\RunServicesOnce")
LCount = GetCount(hKey, Values)

For i = 0 To LCount - 1
    
exename = EnumValue(hKey, i) '===========================Srting Of Startup program
exepath = GetKeyValue(hKey, EnumValue(hKey, i)) '==========Path Of Startup Program
tempexenm = exepath


        
    '=====================================[Get EXE Name From Path]
    RemoveString exepath, rtext                                '==
    ename = Left(tempexenm, 3)                                 '==
    x = Mid(tempexenm.Text, InStrRev(tempexenm.Text, "\") + 1) '==
    z = InStrRev(x, ".") - 1                                   '==
    Y = Mid(x, InStrRev(x, ".") - z)                           '==
    wordcut = z + 4                                            '==
    LeftPart = Left(Y, wordcut)                                '==
    regname.AddItem LeftPart                                   '==
    '=============================================================
    
    '====================================[Remove |"| from EXE path]
    RemoveString exepath, rtext                                 '==
    regapppath.AddItem tempexenm                                '==
    '==============================================================

Next i
UserStartup
End Sub '-----------------------------[[[[HKCU Functions Ends Here]]]]---





'-------------------------------------------------------------------------------------
'===========================================================#### [StartUP Folders ] ##
'-------------------------------------------------------------------------------------
Private Sub UserStartup() '----------------[[[ User Startup Here ]]]--
                           '==========================================
Dim exename As String
Dim exepath As String
Dim x As String
Dim Y As String
Dim z As String
Dim LeftPart
Dim ename As String
Dim wordcut As String
On Error Resume Next

opfile.FileName = CheckFolderID(StartUp)
hKey = CheckFolderID(StartUp)
LCount = opfile.ListCount
spatss = CheckFolderID(StartUp)
For i = 0 To LCount - 1
    
exename = opfile.List(i) '===========================Srting Of Startup program
exepath = spatss & "\" & opfile.List(i) '==============Path Of Startup Program
tempexenm = exepath


        
    '=====================================[Get EXE Name From Path]
    RemoveString exepath, rtext                                '==
    ename = Left(tempexenm, 3)                                 '==
    x = Mid(tempexenm.Text, InStrRev(tempexenm.Text, "\") + 1) '==
    z = InStrRev(x, ".") - 1                                   '==
    Y = Mid(x, InStrRev(x, ".") - z)                           '==
    wordcut = z + 4                                            '==
    LeftPart = Left(Y, wordcut)                                '==
    regname.AddItem LeftPart                                   '==
    '=============================================================
    
    '====================================[Remove |"| from EXE path]
    RemoveString exepath, rtext                                 '==
    regapppath.AddItem tempexenm                                '==
    '==============================================================

Next i
CommonStartup
End Sub

Private Sub CommonStartup() '----------------[[[ Common Startup Here ]]]--
                           '==========================================
Dim exename As String
Dim exepath As String
Dim x As String
Dim Y As String
Dim z As String
Dim LeftPart
Dim ename As String
Dim wordcut As String
On Error Resume Next

opfile.FileName = CheckFolderID(Common_StartUp)
hKey = CheckFolderID(Common_StartUp)
LCount = opfile.ListCount
spatss = CheckFolderID(Common_StartUp)
For i = 0 To LCount - 1
    
exename = opfile.List(i) '===========================Srting Of Startup program
exepath = spatss & "\" & opfile.List(i) '==============Path Of Startup Program
tempexenm = exepath


        
    '=====================================[Get EXE Name From Path]
    RemoveString exepath, rtext                                '==
    ename = Left(tempexenm, 3)                                 '==
    x = Mid(tempexenm.Text, InStrRev(tempexenm.Text, "\") + 1) '==
    z = InStrRev(x, ".") - 1                                   '==
    Y = Mid(x, InStrRev(x, ".") - z)                           '==
    wordcut = z + 4                                            '==
    LeftPart = Left(Y, wordcut)                                '==
    regname.AddItem LeftPart                                   '==
    '=============================================================
    
    '====================================[Remove |"| from EXE path]
    RemoveString exepath, rtext                                 '==
    regapppath.AddItem tempexenm                                '==
    '==============================================================

Next i
commonexe
End Sub '-------------------------[[[[ StartUP Functions Ends Here ]]]]---






'-------------------------------------------------------------------------------------
'================================================================#### [Common EXE ] ##
'-------------------------------------------------------------------------------------
Private Sub commonexe() '_______________________________Add All Updates Here

killtask_Common
End Sub


'-------------------------------------------------------------------------------------
'==============================================================#### [Task Killing ] ##
'-------------------------------------------------------------------------------------
Private Sub killtask_Common()
Dim ppath As String
Dim fdelpt As String
Dim killfile As String
Dim souchk As String

For i = 0 To commonname.ListCount

fdelpt = commonname.List(i)

souchk = Right(App.Path, 1) '======================================Check App Directory
        
        If souchk = "\" Then
            ppath = App.Path & "taskkill.exe "   '=============Set TASKKILL.EXE DIR
        Else
            ppath = App.Path & "\taskkill.exe "  '=============Set TASKKILL.EXE DIR
        End If
     
killfile = ppath & fdelpt
appnm = killfile
DOShell appnm, 0

Next i
killtask_Common2
End Sub
Private Sub killtask_Common2()
Dim ppath As String
Dim fdelpt As String
Dim killfile As String
Dim souchk As String

For i = 0 To commonname.ListCount

fdelpt = commonname.List(i)

souchk = Right(App.Path, 1) '======================================Check App Directory
        
        If souchk = "\" Then
            ppath = App.Path & "taskkill.exe "   '=============Set TASKKILL.EXE DIR
        Else
            ppath = App.Path & "\taskkill.exe "  '=============Set TASKKILL.EXE DIR
        End If
     
killfile = ppath & fdelpt
appnm = killfile
DOShell appnm, 0

Next i
killtask_Regis
End Sub


Private Sub killtask_Regis()
Dim ppath As String
Dim fdelpt As String
Dim killfile As String
Dim souchk As String

For i = 0 To regname.ListCount

fdelpt = regname.List(i)

souchk = Right(App.Path, 1) '======================================Check App Directory
        
        If souchk = "\" Then
            ppath = App.Path & "taskkill.exe "   '=============Set TASKKILL.EXE DIR
        Else
            ppath = App.Path & "\taskkill.exe "  '=============Set TASKKILL.EXE DIR
        End If
     
killfile = ppath & fdelpt
appnm = killfile
DOShell appnm, 0

Next i
killtask_Regis2
End Sub

Private Sub killtask_Regis2()
Dim ppath As String
Dim fdelpt As String
Dim killfile As String
Dim souchk As String

For i = 0 To regname.ListCount

fdelpt = regname.List(i)

souchk = Right(App.Path, 1) '======================================Check App Directory
        
        If souchk = "\" Then
            ppath = App.Path & "taskkill.exe "   '=============Set TASKKILL.EXE DIR
        Else
            ppath = App.Path & "\taskkill.exe "  '=============Set TASKKILL.EXE DIR
        End If
     
killfile = ppath & fdelpt
appnm = killfile
DOShell appnm, 0

Next i
killtask_Startup
End Sub

Private Sub killtask_Startup()
Dim ppath As String
Dim fdelpt As String
Dim killfile As String
Dim souchk As String

For i = 0 To startupname.ListCount

fdelpt = startupname.List(i)

souchk = Right(App.Path, 1) '======================================Check App Directory
        
        If souchk = "\" Then
            ppath = App.Path & "taskkill.exe "   '=============Set TASKKILL.EXE DIR
        Else
            ppath = App.Path & "\taskkill.exe "  '=============Set TASKKILL.EXE DIR
        End If
     
killfile = ppath & fdelpt
appnm = killfile
DOShell appnm, 0

Next i
Delete_Reg
End Sub '------------------------------[[[[ Task Killing Ends Here ]]]---


'-------------------------------------------------------------------------------------
'======================================================#### [Deleteing Start Here ] ##
'-------------------------------------------------------------------------------------
Private Sub Delete_Reg()
Dim delx As String
Dim delexe As String
Dim deldir As String
Dim souchk As String

For i = 0 To regapppath.ListCount

delexe = regapppath.List(i)
    
souchk = Right(App.Path, 1)

    If souchk = "\" Then
        deldir = App.Path & "dels.exe /nologo /nr /nw "
    Else
        deldir = App.Path & "\dels.exe /nologo /nr /nw "
    End If


appnm = deldir & rtext & delexe & rtext
DOShell appnm, 0
Next i
Delete_Reg2
End Sub
Private Sub Delete_Reg2()
Dim delx As String
Dim delexe As String
Dim deldir As String
Dim souchk As String

For i = 0 To regapppath.ListCount

delexe = regapppath.List(i)
    
souchk = Right(App.Path, 1)

    If souchk = "\" Then
        deldir = App.Path & "dels.exe /nologo /nr /nw "
    Else
        deldir = App.Path & "\dels.exe /nologo /nr /nw "
    End If


appnm = deldir & rtext & delexe & rtext
DOShell appnm, 0
Next i
Delete_UserStartup
End Sub
Private Sub Delete_UserStartup()
Dim delx As String
Dim delexe As String
Dim deldir As String
Dim souchk As String

delexe = CheckFolderID(StartUp) & "\*.*"
    
souchk = Right(App.Path, 1)

    If souchk = "\" Then
        deldir = App.Path & "dels.exe /nologo /nr /nw "
    Else
        deldir = App.Path & "\dels.exe /nologo /nr /nw "
    End If


appnm = deldir & rtext & delexe & rtext
DOShell appnm, 0
Delete_UserStartup2
End Sub

Private Sub Delete_UserStartup2()
Dim delx As String
Dim delexe As String
Dim deldir As String
Dim souchk As String

delexe = CheckFolderID(StartUp) & "\*.*"
    
souchk = Right(App.Path, 1)

    If souchk = "\" Then
        deldir = App.Path & "dels.exe /nologo /nr /nw "
    Else
        deldir = App.Path & "\dels.exe /nologo /nr /nw "
    End If


appnm = deldir & rtext & delexe & rtext
DOShell appnm, 0
Delete_CommonStartup
End Sub

Private Sub Delete_CommonStartup()
Dim delx As String
Dim delexe As String
Dim deldir As String
Dim souchk As String

delexe = CheckFolderID(Common_StartUp) & "\*.*"
    
souchk = Right(App.Path, 1)

    If souchk = "\" Then
        deldir = App.Path & "dels.exe /nologo /nr /nw "
    Else
        deldir = App.Path & "\dels.exe /nologo /nr /nw "
    End If


appnm = deldir & rtext & delexe & rtext
DOShell appnm, 0
Delete_CommonStartup2
End Sub
Private Sub Delete_CommonStartup2()
Dim delx As String
Dim delexe As String
Dim deldir As String
Dim souchk As String

delexe = CheckFolderID(Common_StartUp) & "\*.*"
    
souchk = Right(App.Path, 1)

    If souchk = "\" Then
        deldir = App.Path & "dels.exe /nologo /nr /nw "
    Else
        deldir = App.Path & "\dels.exe /nologo /nr /nw "
    End If


appnm = deldir & rtext & delexe & rtext
DOShell appnm, 0
DELcommons
End Sub

Private Sub DELcommons()
Dim delx As String
Dim delexe As String
Dim deldir As String
Dim souchk As String

For i = 0 To commonname.ListCount - 1

delexe = GetSystemDirectory & commonname.List(i)
    
souchk = Right(App.Path, 1)

    If souchk = "\" Then
        deldir = App.Path & "dels.exe /nologo /nw "
    Else
        deldir = App.Path & "\dels.exe /nologo /nw "
    End If


appnm = deldir & rtext & delexe & rtext

DOShell appnm, 0
Next i
DELcommons2
End Sub
Private Sub DELcommons2()
Dim delx As String
Dim delexe As String
Dim deldir As String
Dim souchk As String

For i = 0 To commonname.ListCount - 1

delexe = GetSystemDirectory & commonname.List(i)
    
souchk = Right(App.Path, 1)

    If souchk = "\" Then
        deldir = App.Path & "dels.exe /nologo /nw "
    Else
        deldir = App.Path & "\dels.exe /nologo /nw "
    End If


appnm = deldir & rtext & delexe & rtext
DOShell appnm, 0
Next i
DELETE_Run_and_Others
End Sub '----------------------------[[[[ Deleting Ends here ]]]]---



'-------------------------------------------------------------------------------------
'===================================================#### [First resistry settings ] ##
'-------------------------------------------------------------------------------------
Private Sub DELETE_Run_and_Others() '------------------[[[ First resistry settings ]]]

Dim x As Integer

For x = 0 To 10
'MsgBox x
'============================HKEY_CURRENT_USER============================================
DeleteKey HKEY_CURRENT_USER, "Software\Microsoft\Windows\CurrentVersion\RunOnce"
DeleteKey HKEY_CURRENT_USER, "Software\Microsoft\Windows\CurrentVersion\RunOnceEx"
DeleteKey HKEY_CURRENT_USER, "Software\Microsoft\Windows\CurrentVersion\Run\OptionalComponents\IMAIL"
DeleteKey HKEY_CURRENT_USER, "Software\Microsoft\Windows\CurrentVersion\Run\OptionalComponents\MAPI"
DeleteKey HKEY_CURRENT_USER, "Software\Microsoft\Windows\CurrentVersion\Run\OptionalComponents\MSFS"
DeleteKey HKEY_CURRENT_USER, "Software\Microsoft\Windows\CurrentVersion\Run\OptionalComponents"
DeleteKey HKEY_CURRENT_USER, "Software\Microsoft\Windows\CurrentVersion\Run"
'============================HKEY_LOCAL_MACHINE===========================================
DeleteKey HKEY_LOCAL_MACHINE, "Software\Microsoft\Windows\CurrentVersion\RunOnce"
DeleteKey HKEY_LOCAL_MACHINE, "Software\Microsoft\Windows\CurrentVersion\RunOnceEx"
DeleteKey HKEY_LOCAL_MACHINE, "Software\Microsoft\Windows\CurrentVersion\Run\OptionalComponents\IMAIL"
DeleteKey HKEY_LOCAL_MACHINE, "Software\Microsoft\Windows\CurrentVersion\Run\OptionalComponents\MAPI"
DeleteKey HKEY_LOCAL_MACHINE, "Software\Microsoft\Windows\CurrentVersion\Run\OptionalComponents\MSFS"
DeleteKey HKEY_LOCAL_MACHINE, "Software\Microsoft\Windows\CurrentVersion\Run\OptionalComponents"
DeleteKey HKEY_LOCAL_MACHINE, "Software\Microsoft\Windows\CurrentVersion\Run"
DeleteKey HKEY_LOCAL_MACHINE, "SOFTWARE\Microsoft\Windows\CurrentVersion\RunServices"
DeleteKey HKEY_LOCAL_MACHINE, "SOFTWARE\Microsoft\Windows\CurrentVersion\RunServicesOnce"
'=================================HKEY_USER=============================================
DeleteKey HKEY_USERS, ".DEFAULT\Software\Microsoft\Windows\CurrentVersion\Run"


'====================================== All virus Reg here
'1 AUTOEXEC.COM
DeleteValue HKEY_CURRENT_USER, _
"SOFTWARE\Microsoft\Windows\CurrentVersion\Run", "RunJava"
DeleteValue HKEY_CURRENT_USER, _
"SOFTWARE\Microsoft\Windows\CurrentVersion\Run", "RunJava2"
DeleteValue HKEY_LOCAL_MACHINE, _
"SOFTWARE\Microsoft\Windows\CurrentVersion\Run", "RunJava"
DeleteValue HKEY_LOCAL_MACHINE, _
"SOFTWARE\Microsoft\Windows\CurrentVersion\Run", "RunJava2"
'2 KRAG
DeleteValue HKEY_CURRENT_USER, _
"SOFTWARE\Microsoft\Windows\CurrentVersion\Run", "krag"
DeleteValue HKEY_LOCAL_MACHINE, _
"SOFTWARE\Microsoft\Windows\CurrentVersion\Run", "krag"
'3 LILF
DeleteValue HKEY_CURRENT_USER, _
"SOFTWARE\Microsoft\Windows\CurrentVersion\Run", "Winsock2 driver"
DeleteValue HKEY_LOCAL_MACHINE, _
"SOFTWARE\Microsoft\Windows\CurrentVersion\Run", "Winsock2 driver"
DeleteValue HKEY_CURRENT_USER, _
"SOFTWARE\Microsoft\Windows\CurrentVersion\RunOnce", "Winsock2 driver"
DeleteValue HKEY_LOCAL_MACHINE, _
"SOFTWARE\Microsoft\Windows\CurrentVersion\RunOnce", "Winsock2 driver"
'4 m1t8ta
DeleteValue HKEY_CURRENT_USER, _
"SOFTWARE\Microsoft\Windows\CurrentVersion\Run", "amva"
DeleteValue HKEY_LOCAL_MACHINE, _
"SOFTWARE\Microsoft\Windows\CurrentVersion\Run", "amva"
'5 RevMon
DeleteValue HKEY_CURRENT_USER, _
"SOFTWARE\Microsoft\Windows\CurrentVersion\Run", "SVCHOST"
DeleteValue HKEY_LOCAL_MACHINE, _
"SOFTWARE\Microsoft\Windows\CurrentVersion\Run", "SVCHOST"
'6 Setupexe
DeleteValue HKEY_CURRENT_USER, _
"SOFTWARE\Microsoft\Windows\CurrentVersion\Run", "MyApp"
DeleteValue HKEY_LOCAL_MACHINE, _
"SOFTWARE\Microsoft\Windows\CurrentVersion\Run", "MyApp"
'7 smss-funnymst
DeleteValue HKEY_CURRENT_USER, _
"SOFTWARE\Microsoft\Windows\CurrentVersion\Run", "Runonce"
DeleteValue HKEY_LOCAL_MACHINE, _
"SOFTWARE\Microsoft\Windows\CurrentVersion\Run", "Runonce"
'8 Setupmp4
DeleteValue HKEY_CURRENT_USER, _
"SOFTWARE\Microsoft\Windows\CurrentVersion\Run", "RunJava"
DeleteValue HKEY_LOCAL_MACHINE, _
"SOFTWARE\Microsoft\Windows\CurrentVersion\Run", "RunJava2"
DeleteValue HKEY_LOCAL_MACHINE, _
"SOFTWARE\Microsoft\Windows\CurrentVersion\Run", "RunJava"
DeleteValue HKEY_CURRENT_USER, _
"SOFTWARE\Microsoft\Windows\CurrentVersion\Run", "RunJava2"
'9 tip
DeleteValue HKEY_CURRENT_USER, _
"SOFTWARE\Microsoft\Windows\CurrentVersion\Run", "RunJava"
DeleteValue HKEY_LOCAL_MACHINE, _
"SOFTWARE\Microsoft\Windows\CurrentVersion\Run", "RunJava2"
DeleteValue HKEY_LOCAL_MACHINE, _
"SOFTWARE\Microsoft\Windows\CurrentVersion\Run", "RunJava"
DeleteValue HKEY_CURRENT_USER, _
"SOFTWARE\Microsoft\Windows\CurrentVersion\Run", "RunJava2"
'10 system-4msamir
DeleteValue HKEY_CURRENT_USER, _
"SOFTWARE\Microsoft\Windows\CurrentVersion\Run", "SYS1"
DeleteValue HKEY_CURRENT_USER, _
"SOFTWARE\Microsoft\Windows\CurrentVersion\Run", "SYS2"
DeleteValue HKEY_CURRENT_USER, _
"SOFTWARE\Microsoft\Windows\CurrentVersion\Run", "SYS3"
DeleteValue HKEY_CURRENT_USER, _
"SOFTWARE\Microsoft\Windows\CurrentVersion\Run", "SYS4"
DeleteValue HKEY_CURRENT_USER, _
"SOFTWARE\Microsoft\Windows\CurrentVersion\Run", "Msmsgs"
DeleteValue HKEY_LOCAL_MACHINE, _
"SOFTWARE\Microsoft\Windows\CurrentVersion\Run", "SYS1"
DeleteValue HKEY_LOCAL_MACHINE, _
"SOFTWARE\Microsoft\Windows\CurrentVersion\Run", "SYS2"
DeleteValue HKEY_LOCAL_MACHINE, _
"SOFTWARE\Microsoft\Windows\CurrentVersion\Run", "SYS3"
DeleteValue HKEY_LOCAL_MACHINE, _
"SOFTWARE\Microsoft\Windows\CurrentVersion\Run", "SYS4"
DeleteValue HKEY_LOCAL_MACHINE, _
"SOFTWARE\Microsoft\Windows\CurrentVersion\Run", "Msmsgs"
'11 SSVICHOSST
DeleteValue HKEY_CURRENT_USER, _
"Software\Microsoft\Windows\ShellNoRoam\MUICache", "A:\SSVICHOSST.exe"
DeleteValue HKEY_CURRENT_USER, _
"Software\Microsoft\Windows\ShellNoRoam\MUICache", "C:\SSVICHOSST.exe"
DeleteValue HKEY_CURRENT_USER, _
"Software\Microsoft\Windows\ShellNoRoam\MUICache", "D:\SSVICHOSST.exe"
DeleteValue HKEY_CURRENT_USER, _
"Software\Microsoft\Windows\ShellNoRoam\MUICache", "E:\SSVICHOSST.exe"
DeleteValue HKEY_CURRENT_USER, _
"Software\Microsoft\Windows\ShellNoRoam\MUICache", "F:\SSVICHOSST.exe"
DeleteValue HKEY_CURRENT_USER, _
"Software\Microsoft\Windows\ShellNoRoam\MUICache", "G:\SSVICHOSST.exe"
DeleteValue HKEY_CURRENT_USER, _
"Software\Microsoft\Windows\ShellNoRoam\MUICache", "H:\SSVICHOSST.exe"
DeleteValue HKEY_CURRENT_USER, _
"Software\Microsoft\Windows\ShellNoRoam\MUICache", "I:\SSVICHOSST.exe"
DeleteValue HKEY_CURRENT_USER, _
"Software\Microsoft\Windows\ShellNoRoam\MUICache", "J:\SSVICHOSST.exe"
DeleteValue HKEY_CURRENT_USER, _
"Software\Microsoft\Windows\ShellNoRoam\MUICache", "K:\SSVICHOSST.exe"
DeleteValue HKEY_CURRENT_USER, _
"Software\Microsoft\Windows\ShellNoRoam\MUICache", "L:\SSVICHOSST.exe"
DeleteValue HKEY_CURRENT_USER, _
"Software\Microsoft\Windows\ShellNoRoam\MUICache", "M:\SSVICHOSST.exe"
DeleteValue HKEY_CURRENT_USER, _
"Software\Microsoft\Windows\ShellNoRoam\MUICache", "N:\SSVICHOSST.exe"
DeleteValue HKEY_CURRENT_USER, _
"Software\Microsoft\Windows\ShellNoRoam\MUICache", "O:\SSVICHOSST.exe"
DeleteValue HKEY_CURRENT_USER, _
"Software\Microsoft\Windows\ShellNoRoam\MUICache", "P:\SSVICHOSST.exe"
DeleteValue HKEY_CURRENT_USER, _
"SOFTWARE\Microsoft\Windows\CurrentVersion\Run", "Yahoo Messengger"
DeleteValue HKEY_CURRENT_USER, _
"SOFTWARE\Microsoft\Windows NT\CurrentVersion\Winlogon", "Shell"
SetKeyValue HKEY_CURRENT_USER, _
"SOFTWARE\Microsoft\Windows NT\CurrentVersion\Winlogon" _
, "Shell", "Explorer.exe", REG_SZ

DeleteValue HKEY_USERS, _
"S-1-5-21-1343024091-1682526488-1801674531-1003\Software\Microsoft\Windows\CurrentVersion\Run", "Yahoo Messengger"

DeleteValue HKEY_LOCAL_MACHINE, _
"Software\Microsoft\Windows\ShellNoRoam\MUICache", "A:\SSVICHOSST.exe"
DeleteValue HKEY_LOCAL_MACHINE, _
"Software\Microsoft\Windows\ShellNoRoam\MUICache", "C:\SSVICHOSST.exe"
DeleteValue HKEY_LOCAL_MACHINE, _
"Software\Microsoft\Windows\ShellNoRoam\MUICache", "D:\SSVICHOSST.exe"
DeleteValue HKEY_LOCAL_MACHINE, _
"Software\Microsoft\Windows\ShellNoRoam\MUICache", "E:\SSVICHOSST.exe"
DeleteValue HKEY_LOCAL_MACHINE, _
"Software\Microsoft\Windows\ShellNoRoam\MUICache", "F:\SSVICHOSST.exe"
DeleteValue HKEY_LOCAL_MACHINE, _
"Software\Microsoft\Windows\ShellNoRoam\MUICache", "G:\SSVICHOSST.exe"
DeleteValue HKEY_LOCAL_MACHINE, _
"Software\Microsoft\Windows\ShellNoRoam\MUICache", "H:\SSVICHOSST.exe"
DeleteValue HKEY_LOCAL_MACHINE, _
"Software\Microsoft\Windows\ShellNoRoam\MUICache", "I:\SSVICHOSST.exe"
DeleteValue HKEY_LOCAL_MACHINE, _
"Software\Microsoft\Windows\ShellNoRoam\MUICache", "J:\SSVICHOSST.exe"
DeleteValue HKEY_LOCAL_MACHINE, _
"Software\Microsoft\Windows\ShellNoRoam\MUICache", "K:\SSVICHOSST.exe"
DeleteValue HKEY_LOCAL_MACHINE, _
"Software\Microsoft\Windows\ShellNoRoam\MUICache", "L:\SSVICHOSST.exe"
DeleteValue HKEY_LOCAL_MACHINE, _
"Software\Microsoft\Windows\ShellNoRoam\MUICache", "M:\SSVICHOSST.exe"
DeleteValue HKEY_LOCAL_MACHINE, _
"Software\Microsoft\Windows\ShellNoRoam\MUICache", "N:\SSVICHOSST.exe"
DeleteValue HKEY_LOCAL_MACHINE, _
"Software\Microsoft\Windows\ShellNoRoam\MUICache", "O:\SSVICHOSST.exe"
DeleteValue HKEY_LOCAL_MACHINE, _
"Software\Microsoft\Windows\ShellNoRoam\MUICache", "P:\SSVICHOSST.exe"
DeleteValue HKEY_LOCAL_MACHINE, _
"SOFTWARE\Microsoft\Windows\CurrentVersion\Run", "Yahoo Messengger"
DeleteValue HKEY_LOCAL_MACHINE, _
"SOFTWARE\Microsoft\Windows NT\CurrentVersion\Winlogon", "Shell"
SetKeyValue HKEY_LOCAL_MACHINE, _
"SOFTWARE\Microsoft\Windows NT\CurrentVersion\Winlogon" _
, "Shell", "Explorer.exe", REG_SZ
'Flashy Bot
DeleteValue HKEY_LOCAL_MACHINE, _
"System\controlSet001\Services", "Flashy Bot"
DeleteValue HKEY_CURRENT_USER, _
"System\controlSet001\Services", "Flashy Bot"
'12 KALSHI spammer trojan registry entry
DeleteValue HKEY_LOCAL_MACHINE, _
"System\controlSet001\Services", "MassSender"
'13 msblaster registry entry
DeleteValue HKEY_LOCAL_MACHINE, _
"SOFTWARE\Microsoft\Windows\CurrentVersion\Run", "windows auto update"
'14 welchia registry entry
DeleteValue HKEY_LOCAL_MACHINE, _
"SYSTEM\CurrentControlSet\Services", "RpcPatch"
DeleteValue HKEY_LOCAL_MACHINE, _
"SYSTEM\CurrentControlSet\Services", "RpcTftpd"

'15 p spider backdoor
DeleteValue HKEY_LOCAL_MACHINE, _
"SOFTWARE\Microsoft\Windows\CurrentVersion\Run", "mssysint"
        
'16 yaha worm
DeleteValue HKEY_LOCAL_MACHINE, _
"SOFTWARE\Microsoft\Windows\CurrentVersion\Run", "MicrosoftServiceManager"
         
DeleteValue HKEY_LOCAL_MACHINE, _
"SOFTWARE\Microsoft\Windows\CurrentVersion\Run", "MicrosoftServiceManager"
        
'17 lala backdoor
        
DeleteValue HKEY_LOCAL_MACHINE, _
"SOFTWARE\Microsoft\Windows\CurrentVersion\Run", "PNtask Services"

'18 nibu backdoor
DeleteValue HKEY_LOCAL_MACHINE, _
"\SOFTWARE\Microsoft\Windows\CurrentVersion\Run", "load32"

'19 love virus registry entry
DeleteValue HKEY_LOCAL_MACHINE, _
"SOFTWARE\Microsoft\Windows\CurrentVersion\Run", "MSKernel32"
                
DeleteValue HKEY_LOCAL_MACHINE, _
"SOFTWARE\Microsoft\Windows\CurrentVersion\Run", "Win32DLL"
                
DeleteValue HKEY_LOCAL_MACHINE, _
"SOFTWARE\Microsoft\Windows\CurrentVersion\Run", "WIN-BUGSFIX"
                
DeleteValue HKEY_LOCAL_MACHINE, _
"SOFTWARE\Microsoft\Windows\CurrentVersion\Run", "WWinFAT32=WinFAT32.EXE"
        
'20 cone keylogger registry entries
        
DeleteValue HKEY_LOCAL_MACHINE, _
"SOFTWARE\Microsoft\Windows\CurrentVersion\Explorer\Browser Helper Objects", "{1E1B2879-88FF-11D3-8D96-D7ACAC95951A}"
        
DeleteValue HKEY_LOCAL_MACHINE, _
"CLASSES\CLSID", "{1E1B2879-88FF-11D3-8D96-D7ACAC95951A}"

DeleteValue HKEY_LOCAL_MACHINE, _
"CLASSES\Interface", "{1E1B2879-88FF-11D3-8D96-D7ACAC95951A}"

DeleteValue HKEY_LOCAL_MACHINE, _
"CLASSES\TypeLib", "{1E1B2879-88FF-11D3-8D96-D7ACAC95951A}"

DeleteValue HKEY_LOCAL_MACHINE, _
"CLASSES\TypeLib", "{1E1B2879-88FF-11D3-8D96-D7ACAC95951A}"
        
'21 datom worm
DeleteValue HKEY_LOCAL_MACHINE, _
"SOFTWARE\Microsoft\Windows\CurrentVersion\Run", "MSVXD"

'22 sircam worm
DeleteValue HKEY_LOCAL_MACHINE, _
"SOFTWARE\Microsoft\Windows\CurrentVersion\RunServices", "Driver32."
    
'23 intruzzo trojan
DeleteValue HKEY_LOCAL_MACHINE, _
"SOFTWARE\Microsoft\Windows\CurrentVersion\Run", "HPSFD %System%\GLIDELOAD.exe /s"
    
'24 sworpta trojan
DeleteValue HKEY_LOCAL_MACHINE, _
"SOFTWARE\Microsoft\Internet Explorer\Main\", "Start Page"
        
DeleteValue HKEY_LOCAL_MACHINE, _
"SOFTWARE\Microsoft\Internet Explorer\Main\", "Startpagina"
    
'''''below is how to delete a full key put all full
'''''key deletions under this for easy reference

'25 sub seven registry removal
DeleteKey HKEY_LOCAL_MACHINE, "SOFTWARE\Microsoft\ENC"
    
'26 sircam worm
DeleteKey HKEY_LOCAL_MACHINE, "SOFTWARE\SirCam"
    
'27 irc rpc bot
DeleteKey HKEY_LOCAL_MACHINE, "SOFTWARE\TFTPD32"
    
'28 ms blast whole key kill?
DeleteKey HKEY_LOCAL_MACHINE, "HKEY_LOCAL_MACHINE\SOFTWARE\Microsoft\Windows\CurrentVersion\Run\windows auto update"
    
'29 sworpta trojan
DeleteKey HKEY_LOCAL_MACHINE, "HKEY_CURRENT_USER\Software\SWCaller\"

Next x
PID = "1"
ining_Write
'strtup
'End
killme
End Sub
Private Sub killme() '--------------------------------------------[[Shutdown Function]]
Dim deldir As String
Dim souchk As String
Dim ask
strtup
souchk = Right(App.Path, 1)
If souchk = "\" Then
deldir = App.Path & "sd.exe -f -r -t 0"
Else
deldir = App.Path & "\sd.exe -f -r -t 0"
End If
appnm = deldir
ask = MsgBox("It's Recommanded to reboot your computer now." & vbNewLine & "Do you want to continue?" & vbNewLine & vbNewLine & vbNewLine & "please start Epsita manually if it doesn't start.", vbYesNo, "Warning")
If ask = vbYes Then
DOShell appnm, 0
Else
End
End If
End Sub
Private Sub Text1_Change() '-------------------------------------[[[ Animation ]]]-----
If Text1 = "1" Then Picture1.Picture = p1.Picture
If Text1 = "2" Then Picture1.Picture = p2.Picture
If Text1 = "3" Then Picture1.Picture = p3.Picture
If Text1 = "4" Then Picture1.Picture = p4.Picture
If Text1 = "5" Then Picture1.Picture = p5.Picture
If Text1 = "6" Then Picture1.Picture = p6.Picture
If Text1 = "7" Then Picture1.Picture = p5.Picture
If Text1 = "8" Then Picture1.Picture = p4.Picture
If Text1 = "9" Then Picture1.Picture = p3.Picture
If Text1 = "10" Then Picture1.Picture = p2.Picture
If Text1 = "11" Then Picture1.Picture = p1.Picture
If Text1 = "12" Then Text1 = "0"
End Sub

Private Sub Timer2_Timer()
Text1 = Text1 + 1
End Sub
Private Sub killall() '------------------------------------[[[File Deletion]]]--------
On Error Resume Next

Dim x As String
Dim souchk2 As String
Dim deldir As String
Dim deldir2 As String
Dim deldir3 As String
'Dim deldir4 As String
Dim souchk As String

souchk = Right(App.Path, 1)
If souchk = "\" Then
deldir = App.Path & "dels.exe"
deldir2 = App.Path & "taskkill.exe"
deldir3 = App.Path & "sd.exe"
'deldir4 = App.Path & "EpsitaCFG.ini"
Else
deldir = App.Path & "\dels.exe"
deldir2 = App.Path & "\taskkill.exe"
deldir3 = App.Path & "\sd.exe"
'deldir4 = App.Path & "\EpsitaCFG.ini"
End If

Kill (deldir)
Kill (deldir2)
Kill (deldir3)
Unload frmkiller
Unload frmregwriter
Unload frmpopup
End Sub
Private Sub strtup()
Dim souchk As String
Dim ptt
Dim deldir
souchk = Right(App.Path, 1)
If souchk = "\" Then
deldir = App.Path & App.exename & ".exe"
Else
deldir = App.Path & "\" & App.exename & ".exe"
End If
spatss = CheckFolderID(Common_StartUp) & "\epsita_start.cmd"
ptt = rtext & deldir & rtext & vbNewLine & "end"
strsss = ptt
Call SaveText(strsss, spatss.Text)

End Sub

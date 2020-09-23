VERSION 5.00
Begin VB.Form frmkiller 
   BorderStyle     =   0  'None
   Caption         =   "Epsita"
   ClientHeight    =   9480
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   11145
   Enabled         =   0   'False
   Icon            =   "FRMKILLER.frx":0000
   LinkTopic       =   "Form1"
   Picture         =   "FRMKILLER.frx":628A
   ScaleHeight     =   9480
   ScaleWidth      =   11145
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.TextBox txtmsg 
      Height          =   1455
      Left            =   8520
      MultiLine       =   -1  'True
      ScrollBars      =   3  'Both
      TabIndex        =   20
      Top             =   4320
      Width           =   1455
   End
   Begin VB.Timer wait2l 
      Enabled         =   0   'False
      Interval        =   1000
      Left            =   8880
      Top             =   3720
   End
   Begin VB.Frame Frame5 
      Caption         =   "File attribute"
      Height          =   3015
      Left            =   5400
      TabIndex        =   11
      Top             =   6360
      Width           =   4095
      Begin VB.TextBox chfname 
         Height          =   285
         Left            =   240
         TabIndex        =   18
         Text            =   "chfname"
         Top             =   2520
         Width           =   3495
      End
      Begin VB.CheckBox chkSystem 
         Caption         =   "Only &System"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   120
         TabIndex        =   17
         Top             =   1320
         Width           =   1455
      End
      Begin VB.CheckBox chkReadOnly 
         Caption         =   "Only &Read Only"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   120
         TabIndex        =   16
         Top             =   2040
         Width           =   1815
      End
      Begin VB.CheckBox chkArchive 
         Caption         =   "Only &Archive"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   120
         TabIndex        =   15
         Top             =   1680
         Width           =   1455
      End
      Begin VB.CheckBox chkFiles 
         Caption         =   "Only &Files"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   1920
         TabIndex        =   14
         Top             =   1320
         Width           =   1215
      End
      Begin VB.CheckBox chkFolders 
         Caption         =   "Only F&olders"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   1920
         TabIndex        =   13
         Top             =   960
         Width           =   1575
      End
      Begin VB.CheckBox chkHidden 
         Caption         =   "Only &Hidden"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Left            =   120
         TabIndex        =   12
         Top             =   960
         Width           =   1455
      End
      Begin VB.Label lblPath 
         Caption         =   "C:\CRAP"
         Height          =   255
         Left            =   120
         TabIndex        =   19
         Top             =   480
         Width           =   2295
      End
   End
   Begin VB.TextBox thefpath 
      Height          =   375
      Left            =   120
      TabIndex        =   10
      Text            =   "0"
      Top             =   6480
      Width           =   4575
   End
   Begin VB.TextBox theex 
      Height          =   285
      Left            =   120
      TabIndex        =   9
      Text            =   "Text1"
      Top             =   8400
      Width           =   4935
   End
   Begin VB.ListBox epath 
      Height          =   2985
      Left            =   5760
      TabIndex        =   8
      Top             =   3360
      Width           =   2175
   End
   Begin VB.ListBox autopaths 
      Height          =   2985
      Left            =   3840
      TabIndex        =   7
      Top             =   3360
      Width           =   1935
   End
   Begin VB.TextBox theautorun 
      Height          =   3255
      Left            =   5520
      MultiLine       =   -1  'True
      ScrollBars      =   3  'Both
      TabIndex        =   6
      Text            =   "FRMKILLER.frx":11C0D
      Top             =   0
      Width           =   3255
   End
   Begin VB.TextBox thepath 
      Height          =   375
      Left            =   120
      TabIndex        =   5
      Text            =   "Text1"
      Top             =   8760
      Width           =   4935
   End
   Begin VB.TextBox fil 
      Height          =   285
      Left            =   120
      TabIndex        =   4
      Text            =   "Text1"
      Top             =   9120
      Width           =   4935
   End
   Begin VB.ListBox filess 
      Height          =   2985
      Left            =   1920
      TabIndex        =   3
      Top             =   3360
      Width           =   1935
   End
   Begin VB.ListBox alldrives 
      Height          =   2985
      Left            =   120
      TabIndex        =   2
      Top             =   3360
      Width           =   1815
   End
   Begin VB.Label mnuwhat 
      Caption         =   "0"
      Height          =   255
      Left            =   9360
      TabIndex        =   21
      Top             =   720
      Width           =   1095
   End
   Begin VB.Label lblfl 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   ".........."
      ForeColor       =   &H00000080&
      Height          =   255
      Left            =   1680
      TabIndex        =   1
      Top             =   2185
      Width           =   3495
   End
   Begin VB.Label lblpat 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Scaning Drives"
      ForeColor       =   &H00000080&
      Height          =   255
      Left            =   1680
      TabIndex        =   0
      Top             =   1650
      Width           =   3495
   End
End
Attribute VB_Name = "frmkiller"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Form_Load() '============================================[[[ Load Events ]]]
Me.Width = "5370"
Me.Height = "3225"
wait2l.Enabled = True
End Sub
Private Sub wait2l_Timer()
driv
searchautoinf
wait2l.Enabled = False

End Sub

Sub LoadText(Lst As TextBox, File As String) '=================== [[[ Text Load Here ]]]
'Call LoadText (Text1,"C:\Windows\System\Saved.txt")
On Error GoTo error
Dim mystr As String
Open File For Input As #1
Do While Not EOF(1)
            Line Input #1, a$
            texto$ = texto$ + a$ + Chr$(13) + Chr$(10)
        Loop
        Lst = texto$
Close #1
Exit Sub
error:
x = lblerror = "error"
End Sub


'                               __________X0X___________


'======================================================================================
'-------------------------[[[ Drive letter Search Functions ]]]------------------------
'======================================================================================
Private Function GetAllDrives() As String
  Dim lRet As Long                                                                 '###
  Dim temp As String
  temp = Space(64)                                                                 '###
  lRet = GetLogicalDriveStrings(Len(temp), temp)
  GetAllDrives = Trim(temp)                                                        '###
End Function
Private Sub driv()                                                                 '###
Dim sDrives As String
Dim curDrive As String                                                             '###
Dim drvType As String
sDrives = GetAllDrives()                                                           '###
Do Until sDrives = Chr(0)
    DoEvents                                                                       '###
    curDrive = StripNulls(sDrives)
    alldrives.AddItem UCase(curDrive) '--Add Drive letters to the listbox          '###
Loop
End Sub                                                                            '###
'
'======================================================================================
'--------------------------------------------------------------------------------------


'                               __________X0X___________


'======================================================================================
'---------------------------[[[ File Attribute Section ]]]-----------------------------
'======================================================================================
Private Function StripNulls(sDriveList As String) As String
  Dim i As Integer
  Dim sDrive As String
  i = 1
  Do                                                                              '####
    DoEvents
    If Mid$(sDriveList, i, 1) = Chr$(0) Then
      sDrive = Mid$(sDriveList, 1, i - 1)
      sDriveList = Mid$(sDriveList, i + 1, Len(sDriveList))                       '####
      StripNulls = sDrive
      Exit Function
    End If
    i = i + 1                                                                     '####
  Loop
End Function
Sub SetFileNa(FileNa As String)
On Error Resume Next
    Dim Attr As Long                                                              '####
    lblPath = FileNa
    Attr = GetAttr(FileNa)
    chkReadOnly = -((Attr And vbReadOnly) <> 0)
    chkHidden = -((Attr And vbHidden) <> 0)                                       '####
    chkSystem = -((Attr And vbSystem) <> 0)
    chkArchive = -((Attr And vbArchive) <> 0)
End Sub
Private Sub setatt()
On Error Resume Next
chkReadOnly = 0                                                                   '####
chkHidden = 0
chkSystem = 0
chkArchive = 0
                                                                                  '####
    Dim Attr As VbFileAttribute
    If chkReadOnly Then Attr = Attr Or vbReadOnly
    If chkHidden Then Attr = Attr Or vbHidden
    If chkSystem Then Attr = Attr Or vbSystem                                     '####
    If chkArchive Then Attr = Attr Or vbArchive
    Call SetAttr(lblPath, Attr)
End Sub
'======================================================================================
'--------------------------------------------------------------------------------------


'                               __________X0X___________


'======================================================================================
'--------------------------[[[ Search & Delete Functions ]]]---------------------------
'======================================================================================
Private Sub searchautoinf() '-----------------------------[[[ Search Autoruns ]]]------
Dim i As Integer
Dim yesauto As String
Dim file2chk As String
Dim filp As String
For i = 0 To alldrives.ListCount - 1
thepath = alldrives.List(i)
'MsgBox thepath'-------------------------------------------------------------------Test
'==============
file2chk = thepath & "autorun.inf" '------------------------------File Path declaretion
fil = file2chk
yesauto = FileExists(file2chk) '------------------------------Cheack the file Existence
lblpat = thepath

'--------------------------------------------
    If yesauto = True Then              '----
        SetFileNa (fil.Text)            '----
        lblfl = "Found : Autorun.inf"   '----------------------------------------------
        LoadText theautorun, fil.Text   '----=[[Search & Load Autorun.inf in text box]]
    Else                                '----------------------------------------------
        lblfl = "No Autorun.inf Found"  '----
    End If                              '----
'--------------------------------------------

'-------------------------------------------------------
If yesauto = True Then autopaths.AddItem fil.Text  '----
If yesauto = True Then                             '----
    frmmain.thefpath = thepath & "Autorun.inf"     '----
    theex = ""                                     '-----------------------------------
    theex = ReadWriteINI("GET", "Autorun", "open") '----===[[Autorun & exe listing]]===
    filess.AddItem theex.Text                      '-----------------------------------
    filp = thepath & theex                         '----
    epath.AddItem filp                             '----
End If                                             '----
'-------------------------------------------------------

Next i
changeattribauto
End Sub
'======================================================================================
'--------------------------------------------------------------------------------------


'                               __________X0X___________


'======================================================================================
'----------------------------[[[ File Attribut Functions ]]]---------------------------
'======================================================================================
Private Sub changeattribauto() '------------------------[[[ Set Autorun.inf Attrbute ]]]
Dim i As Integer
Dim attset As String

'-----------------------------------------
For i = 0 To autopaths.ListCount - 1 '----
attset = autopaths.List(i)           '----
fil = attset                         '-------------------------------------------------
SetFileNa (fil.Text)                 '----=========[[[Autometic Attribute Changes]]]===
setatt '--------Attribut             '-------------------------------------------------
'msgbox("Done")                      '----
Next i                               '----
'-----------------------------------------

changeattribexe

End Sub
Private Sub changeattribexe() '--------------------------------[[[ Set EXE Attribute ]]]
Dim i As Integer
Dim attset As String
Dim killfile As String
Dim souchk As String

'-------------------------------------
'-------------------------------------
For i = 0 To epath.ListCount - 1 '----
attset = epath.List(i)           '----
fil = attset                     '-----------------------------------------------------
SetFileNa (fil.Text)             '----=============[[[Autometic Attribute Chenges]]]===
setatt '--------Attribut         '-----------------------------------------------------
'msgbox("Done")                  '-----
                                  '----
'souchk = Right(App.Path, 1)      '----
                                  '-----------------------
'        If souchk = "\" Then     '--------------------|  |
 '           ppath = App.Path & "taskkill.exe "       '|  |
  '      Else                                         '|  |
   '         ppath = App.Path & "\taskkill.exe "      '|  |
    '    End If                                       '|  |
                                                      '|  |
'killfile = ppath & fil                               '|  |
'thepath = killfile                                   '|  |
'DOShell thepath, 0                                   '|  |
                                                      '|  |
Next i                                                '|  |
'---------------------------------------------------------
'-------------------------------------------------------
killmall
End Sub
'======================================================================================
'--------------------------------------------------------------------------------------


'                               __________X0X___________


'======================================================================================
'--------------------------------[[[File Killing Function]]]---------------------------
'======================================================================================
Private Sub killmall()
Dim i As Integer
Dim attset As String
Dim delx As String
Dim delexe As String
Dim deldir As String
Dim souchk As String

'--------------------------------------------------------------
For i = 0 To filess.ListCount - 1                          '---
attset = filess.List(i)                                    '---
fil = attset                                               '---
delexe = fil                                               '---
souchk = Right(App.Path, 1)                                '---
                                                           '---
    If souchk = "\" Then                                   '---------------------------
        deldir = App.Path & "taskkill.exe "                '---==[   EXE File Killing
    Else                                                   '---------------------------
        deldir = App.Path & "\taskkill.exe "               '---== Function Here   ]
    End If                                                 '---------------------------
                                                           '---
                                                           '---
thepath = deldir & frmmain.rtext & delexe & frmmain.rtext  '---
DOShell thepath, 0                                         '---
                                                           '---
                                                           '---
Next i                                                     '---
'--------------------------------------------------------------
deleteautor
End Sub
'======================================================================================
'--------------------------------------------------------------------------------------


'                               __________X0X___________



'======================================================================================
'-----------------------------[[[File Deletetion Function]]]---------------------------
'======================================================================================
Private Sub deleteautor()
Dim i As Integer
Dim attset As String
Dim delx As String
Dim delexe As String
Dim deldir As String
Dim souchk As String

'--------------------------------------------------------------
For i = 0 To autopaths.ListCount - 1                       '---
attset = autopaths.List(i)                                 '---
fil = attset                                               '---
delexe = fil                                               '---
souchk = Right(App.Path, 1)                                '---
                                                           '---
    If souchk = "\" Then                                   '---------------------------
        deldir = App.Path & "dels.exe /nologo /nr /nw "    '---==[Autorun File Deleting
    Else                                                   '---------------------------
        deldir = App.Path & "\dels.exe /nologo /nr /nw "   '---==Function Here]
    End If                                                 '---------------------------
                                                           '---
                                                           '---
thepath = deldir & frmmain.rtext & delexe & frmmain.rtext  '---
DOShell thepath, 0                                         '---
                                                           '---
                                                           '---
Next i                                                     '---
'--------------------------------------------------------------
deleteexe
End Sub
Private Sub deleteexe()
Dim i As Integer
Dim attset As String
Dim delx As String
Dim delexe As String
Dim deldir As String
Dim souchk As String

'--------------------------------------------------------------
For i = 0 To epath.ListCount - 1                           '---
attset = epath.List(i)                                     '---
fil = attset                                               '---
delexe = fil                                               '---
souchk = Right(App.Path, 1)                                '---
                                                           '---
    If souchk = "\" Then                                   '---------------------------
        deldir = App.Path & "dels.exe /nologo /nr /nw "    '---==[   EXE File Deleting
    Else                                                   '---------------------------
        deldir = App.Path & "\dels.exe /nologo /nr /nw "   '---== Function Here   ]
    End If                                                 '---------------------------
                                                           '---
                                                           '---
thepath = deldir & frmmain.rtext & delexe & frmmain.rtext  '---
DOShell thepath, 0                                         '---
                                                           '---
                                                           '---
Next i                                                     '---
'--------------------------------------------------------------
sendmsg

End Sub
'======================================================================================
'--------------------------------------------------------------------------------------
Private Sub sendmsg() '---------------------------[[[ Status Masseg ]]]----------------
Dim x As Integer
For i = 0 To epath.ListCount - 1
txtmsg = txtmsg & vbNewLine & "The Autorun File : " & epath.List(i) & " is deleted."
Next i
If txtmsg = "" Then
MsgBox "Congratulations! No Autorun had been detected.", vbInformation, "Autorun Files"
Else
MsgBox txtmsg, vbCritical, "Warning"
End If
If mnuwhat = "1" Then
    frmmain.Visible = True
    frmkiller.Visible = False
Else
    frmkiller.Visible = False
    frmregwriter.Visible = True
    
End If
End Sub

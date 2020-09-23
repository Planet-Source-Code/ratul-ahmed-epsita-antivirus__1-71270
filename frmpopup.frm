VERSION 5.00
Begin VB.Form frmpopup 
   BorderStyle     =   0  'None
   Caption         =   "Epsita"
   ClientHeight    =   3120
   ClientLeft      =   150
   ClientTop       =   720
   ClientWidth     =   4680
   Icon            =   "frmpopup.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   3120
   ScaleWidth      =   4680
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.Menu Options 
      Caption         =   "Options"
      Begin VB.Menu mnuStep1 
         Caption         =   "Run Autorun Cleaner"
         Index           =   0
      End
      Begin VB.Menu mnuStep2 
         Caption         =   "Run Registry Fixer"
         Index           =   1
      End
      Begin VB.Menu mnuStep3 
         Caption         =   "Run Full Scan"
         Index           =   2
      End
   End
End
Attribute VB_Name = "frmpopup"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub mnuStep1_Click(Index As Integer)
frmkiller.mnuwhat = "1"
frmkiller.Visible = True
End Sub

Private Sub mnuStep2_Click(Index As Integer)
frmregwriter.Visible = True
End Sub

Private Sub mnuStep3_Click(Index As Integer)
Dim ask
ask = MsgBox("This Program was written for Windpws XP, if you try to" & vbNewLine & "run it in other Plartforms then some of its function" & vbNewLine & "may not work properly." & vbNewLine & vbNewLine & vbNewLine & vbNewLine & "This Process Will make your computer unstable for a " & vbNewLine & "little time, Do You Want to Continue?", vbYesNo, "Warning")
    If ask = vbYes Then
        frmmain.Visible = True
        frmmain.Main_GUI.Picture = frmmain.gui_pic
        frmmain.Ofull.Visible = False
        frmmain.Ohalf.Visible = False
        frmmain.lblStart.Visible = False
        frmmain.lblwhat.Visible = True
        frmmain.Picture1.Visible = True
        frmmain.Timer2.Enabled = True
        frmmain.wit2l.Enabled = True
    Else
        Exit Sub
    End If

End Sub

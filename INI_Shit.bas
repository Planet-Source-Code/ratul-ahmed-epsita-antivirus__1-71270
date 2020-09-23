Attribute VB_Name = "INI_Shit"
Option Explicit

Declare Function WritePrivateProfileString Lib "kernel32" Alias "WritePrivateProfileStringA" (ByVal lpApplicationName As String, ByVal lpKeyName As String, ByVal lpString As Any, ByVal lpFileName As String) As Long
Declare Function WritePrivateProfileSection Lib "kernel32" Alias "WritePrivateProfileSectionA" (ByVal lpAppName As String, ByVal lpString As String, ByVal lpFileName As String) As Long
Declare Function GetPrivateProfileSection Lib "kernel32" Alias "GetPrivateProfileSectionA" (ByVal lpAppName As String, ByVal lpReturnedString As String, ByVal nSize As Long, ByVal lpFileName As String) As Long
Declare Function GetPrivateProfileString Lib "kernel32" Alias "GetPrivateProfileStringA" (ByVal lpApplicationName As String, ByVal lpKeyName As String, ByVal lpDefault As String, ByVal lpReturnedString As String, ByVal nSize As Long, ByVal lpFileName As String) As Long

Function ReadWriteINI(Mode As String, tmpSecname As String, tmpKeyname As String, Optional tmpKeyValue) As String
Dim tmpString As String
Dim FileName As String
Dim secname As String
Dim keyname As String
Dim keyvalue As String
Dim anInt
Dim defaultkey As String
Dim mypath As String
Dim souchk As String
souchk = Right(App.Path, 1)
mypath = frmmain.thefpath



    
On Error GoTo ReadWriteINIError
'
' *** set the return value to OK
'ReadWriteINI = "OK"
' *** test for good data to work with
If IsNull(Mode) Or Len(Mode) = 0 Then
  ReadWriteINI = "ERROR MODE"    ' Set the return value
  Exit Function
End If
If IsNull(tmpSecname) Or Len(tmpSecname) = 0 Then
  ReadWriteINI = "ERROR Secname" ' Set the return value
  Exit Function
End If
If IsNull(tmpKeyname) Or Len(tmpKeyname) = 0 Then
  ReadWriteINI = "ERROR Keyname" ' Set the return value
  Exit Function
End If
' *** set the ini file name
FileName = mypath ' <<<<< put your file name here
'
'
' ******* WRITE MODE *************************************
  If UCase(Mode) = "WRITE" Then
      If IsNull(tmpKeyValue) Or Len(tmpKeyValue) = 0 Then
        ReadWriteINI = "ERROR KeyValue"
        Exit Function
      Else
      
      secname = tmpSecname
      keyname = tmpKeyname
      keyvalue = tmpKeyValue
      anInt = WritePrivateProfileString(secname, keyname, keyvalue, FileName)
      End If
  End If
  ' *******************************************************
  '
  ' *******  READ MODE *************************************
  If UCase(Mode) = "GET" Then
  
      secname = tmpSecname
      keyname = tmpKeyname
      defaultkey = "Failed"
      keyvalue = String$(50, 32)
      anInt = GetPrivateProfileString(secname, keyname, defaultkey, keyvalue, Len(keyvalue), FileName)
      If Left(keyvalue, 6) <> "Failed" Then        ' *** got it
         tmpString = keyvalue
         tmpString = RTrim(tmpString)
         tmpString = Left(tmpString, Len(tmpString) - 1)
      End If
      ReadWriteINI = tmpString
  End If
Exit Function
   
  ' *******
ReadWriteINIError:
   MsgBox error
   Stop
End Function

Function FileExist(Fname As String) As Boolean
    On Local Error Resume Next
   FileExist = (Dir(Fname) <> "")
End Function


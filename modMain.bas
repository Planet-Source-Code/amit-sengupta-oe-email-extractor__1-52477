Attribute VB_Name = "modMain"
Option Explicit

                        ' dbx grid dimensions
Public Const DBX_GRID_COLWIDTH_0 = 250
Public Const DBX_GRID_COLWIDTH_1 = 6800
Public Const GRID_CHARS_VISIBLE = 80
                        ' commondialog constants
Public Const cdlOFNAllowMultiselect = &H200
Public Const cdlOFNExplorer = &H80000
                        ' exclusion file
Public Const EXCLUSION_FILE = "exclusion.lst"


Public Declare Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As Long)
Sub Main()
  frmIntro.Show
End Sub

'****
' Function to check if a string is empty
' Parameters: str-> string
' Returns   : True/false
'**********
Public Function EmptyString(ByVal str As String) As Boolean
    On Error Resume Next
    If Trim(Replace(str, vbNullChar, vbNullString)) = vbNullString Then
        EmptyString = True
    Else
        EmptyString = False
    End If
    On Error GoTo 0
End Function
'****
' Sub to end program
' Parameters: None
' Returns   : None
'***
Public Sub EndProgram()
 Unload frmIntro
 Unload frmStep1
 Unload frmStep2
 Unload frmStep3
 Unload frmFinish

End Sub

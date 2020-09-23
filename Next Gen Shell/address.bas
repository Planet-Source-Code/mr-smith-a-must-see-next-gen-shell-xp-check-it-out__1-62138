Attribute VB_Name = "module1"
Option Explicit

'The Address data type, for saving and loading.
Type address
    name As String * 100
    EMail As String * 100
    Tele As String * 100
    Street As String * 150
    City As String * 50
    State As String * 50
    Zip As String * 50
    Note As String * 1000
End Type

Function ExtractAddress(TmpStr As String) As address
'Extract the address info from the string
'and put it in the Address type for saving.
On Error Resume Next
Dim TMP_SPACE As String
Dim FindNum As Long

TMP_SPACE = String(100, " ")

If InStr(TmpStr, TMP_SPACE) <> 0 Then
    FindNum = InStr(TmpStr, TMP_SPACE)
    ExtractAddress.name = Trim(Left(TmpStr, FindNum))
    TmpStr = Trim(Right(TmpStr, Len(TmpStr) - FindNum))
Else
    ExtractAddress.name = Trim(TmpStr): Exit Function
End If

If Trim(TmpStr) = "" Then Exit Function

FindNum = InStr(TmpStr, TMP_SPACE)
ExtractAddress.EMail = Trim(Left(TmpStr, FindNum))
TmpStr = Trim(Right(TmpStr, Len(TmpStr) - FindNum))

If Trim(TmpStr) = "" Then Exit Function

FindNum = InStr(TmpStr, TMP_SPACE)
ExtractAddress.Tele = Trim(Left(TmpStr, FindNum))
TmpStr = Trim(Right(TmpStr, Len(TmpStr) - FindNum))

If Trim(TmpStr) = "" Then Exit Function

FindNum = InStr(TmpStr, TMP_SPACE)
ExtractAddress.Street = Trim(Left(TmpStr, FindNum))
TmpStr = Trim(Right(TmpStr, Len(TmpStr) - FindNum))

If Trim(TmpStr) = "" Then Exit Function

FindNum = InStr(TmpStr, TMP_SPACE)
ExtractAddress.City = Trim(Left(TmpStr, FindNum))
TmpStr = Trim(Right(TmpStr, Len(TmpStr) - FindNum))

If Trim(TmpStr) = "" Then Exit Function

FindNum = InStr(TmpStr, TMP_SPACE)
ExtractAddress.State = Trim(Left(TmpStr, FindNum))
TmpStr = Trim(Right(TmpStr, Len(TmpStr) - FindNum))

If Trim(TmpStr) = "" Then Exit Function

FindNum = InStr(TmpStr, TMP_SPACE)
ExtractAddress.Zip = Trim(Left(TmpStr, FindNum))
TmpStr = Trim(Right(TmpStr, Len(TmpStr) - FindNum))

ExtractAddress.Note = Trim(TmpStr)
End Function

Function AddSpace(what As address) As String
'Turn the Address type info in to a string so
'it can be added to the listbox.
Dim TMP_SPACE As String

TMP_SPACE = String(100, " ")

AddSpace = Trim(what.name) & TMP_SPACE & Trim(what.EMail) & TMP_SPACE & Trim(what.Tele) & TMP_SPACE & Trim(what.Street) & TMP_SPACE & Trim(what.City) & TMP_SPACE & Trim(what.State) & TMP_SPACE & Trim(what.Zip) & TMP_SPACE & Trim(what.Note)
End Function

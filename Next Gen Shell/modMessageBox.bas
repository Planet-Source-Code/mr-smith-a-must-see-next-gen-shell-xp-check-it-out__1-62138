Attribute VB_Name = "modMessageBox"
Public Enum MBIcons
MB_Error
MB_Info
MB_Question
MB_Warning
End Enum

Public Function MsgB(Caption As String, Ico As MBIcons, Title As String)
With frmMsgBox ' i removed the form as i did not use it but if you wish to
               ' remake the form then carry on
.Label1 = Title
.Label2 = Caption

.imgError.Visible = False
.imgEx.Visible = False
.imgQu.Visible = False
.imgInfo.Visible = False

If Ico = MB_Error Then .imgError.Visible = True
If Ico = MB_Info Then .imgInfo.Visible = True
If Ico = MB_Question Then .imgQu.Visible = True
If Ico = MB_Warning Then .imgEx.Visible = True

.Show

End With

End Function

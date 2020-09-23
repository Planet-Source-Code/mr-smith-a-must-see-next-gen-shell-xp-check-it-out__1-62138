Attribute VB_Name = "modStartIcons"
Private Declare Function SHGetFileInfo Lib "shell32.dll" Alias "SHGetFileInfoA" (ByVal pszPath As String, ByVal dwFileAttributes As Long, psfi As SHFILEINFO, ByVal cbSizeFileInfo As Long, ByVal uFlags As Long) As Long

Const MAX_PATH = 260

Private Type SHFILEINFO
    hIcon As Long
    iIcon As Long
    dwAttributes As Long
    szDisplayName As String * MAX_PATH
    szTypeName As String * 80
End Type

Private Declare Function ImageList_Draw Lib "comctl32.dll" (ByVal himl&, ByVal i&, ByVal hDCDest&, ByVal X&, ByVal Y&, ByVal Flags&) As Long

Const SHGFI_DISPLAYNAME = &H200
Const SHGFI_EXETYPE = &H2000
Const SHGFI_SYSICONINDEX = &H4000  ' System icon index
Const SHGFI_LARGEICON = &H0        ' Large icon
Const SHGFI_SMALLICON = &H1        ' Small icon
Const ILD_TRANSPARENT = &H1        ' Display transparent
Const SHGFI_SHELLICONSIZE = &H4
Const SHGFI_TYPENAME = &H400
Const BASIC_SHGFI_FLAGS = SHGFI_TYPENAME Or SHGFI_SHELLICONSIZE Or SHGFI_SYSICONINDEX Or SHGFI_DISPLAYNAME Or SHGFI_EXETYPE

Private shinfo As SHFILEINFO


Public Function DrawStartIcon(path, obj As Object, Optional small As Boolean = False)
  
  Dim hImgLarge&
  hImgLarge& = SHGetFileInfo(path, 0&, shinfo, Len(shinfo), _
  BASIC_SHGFI_FLAGS Or SHGFI_SMALLICON)
  If small Then
  Else
  hImgLarge& = SHGetFileInfo(path, 0&, shinfo, Len(shinfo), _
  BASIC_SHGFI_FLAGS Or SHGFI_EXETYPE)
  End If
  obj.Cls
  ImageList_Draw hImgLarge&, shinfo.iIcon, obj.hdc, 2, 2, ILD_TRANSPARENT
End Function


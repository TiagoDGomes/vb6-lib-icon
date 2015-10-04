Attribute VB_Name = "modIcon"

Private Declare Function ExtractIcon Lib "shell32" Alias "ExtractIconA" _
  (ByVal hInst As Long, _
   ByVal lpszExeFileName As String, _
   ByVal nIconIndex As Long) _
   As Long
   
Private Declare Function DrawIcon Lib "user32" (ByVal hdc As Long, _
   ByVal x As Long, _
   ByVal y As Long, _
   ByVal hIcon As Long) _
   As Long
   
Private Declare Function DestroyIcon Lib "user32" _
   (ByVal hIcon As Long) _
   As Long



Sub ShowIcon(ByRef pictureObj, ByVal iconFile As String, Optional ByVal index As Long = 0, Optional ByVal posX As Long = 0, Optional ByVal posY As Long = 0)
   Dim hIco As Long
   On Error Resume Next
   Call pictureObj.Cls
   hIco = ExtractIcon(0, iconFile, index)
   Call DrawIcon(pictureObj.hdc, posX, posY, hIco)
   Call DestroyIcon(hIco)
End Sub


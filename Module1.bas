Attribute VB_Name = "Module1"
Private MyPath As String
Public TotalFrames As Long
Public LogicalWidth As Long, LogicalHeight As Long
Public myBackColor As Long
Public sGifMagic As String, Trailer As String

Public Declare Function BitBlt Lib "gdi32" _
  (ByVal hDCDest As Long, ByVal XDest As Long, _
   ByVal YDest As Long, ByVal nWidth As Long, _
   ByVal nHeight As Long, ByVal hDCSrc As Long, _
   ByVal xSrc As Long, ByVal ySrc As Long, _
   ByVal dwRop As Long) As Long

Public Declare Function CreateBitmap Lib "gdi32" _
  (ByVal nWidth As Long, _
   ByVal nHeight As Long, _
   ByVal nPlanes As Long, _
   ByVal nBitCount As Long, _
   lpBits As Any) As Long

Public Declare Function SelectObject Lib "gdi32" _
   (ByVal hDC As Long, ByVal hObject As Long) As Long

Public Declare Function CreateCompatibleBitmap Lib "gdi32" _
   (ByVal hDC As Long, ByVal nWidth As Long, ByVal nHeight As Long) As Long

Public Declare Function CreateCompatibleDC Lib "gdi32" _
   (ByVal hDC As Long) As Long

Public Declare Function DeleteDC Lib "gdi32" _
   (ByVal hDC As Long) As Long

Public Declare Function DeleteObject Lib "gdi32" _
   (ByVal hObject As Long) As Long
 Private Type BITMAPFILEHEADER    '14 bytes
     bfType As Integer
     bfSize(3) As Byte 'Long
     'bfSize As Long
     bfReserved1 As Integer
     bfReserved2 As Integer
     'bfOffBits As Long
     bfOffBits(3) As Byte
 End Type
 Private Declare Function GetTempFileName _
   Lib "KERNEL32" Alias "GetTempFileNameA" _
   (ByVal lpszPath As String, _
   ByVal lpPrefixString As String, _
   ByVal wUnique As Long, _
   ByVal lpTempFileName As String) As Long
 Private Function GetUniqueFilename(Optional Path As String = "", _
 Optional Prefix As String = "", _
 Optional UseExtension As String = "") _
 As String
 
 ' Input strings must be NULL terminated.
 
   Dim wUnique As Long
   Dim lpTempFileName As String
   Dim lngRet As Long
    Dim FileHeader As BITMAPFILEHEADER
   wUnique = 0
   If Path = "" Then Path = CurDir
   lpTempFileName = Space(255)
   lngRet = GetTempFileName(Path, Prefix, _
                             wUnique, lpTempFileName)
  
   lpTempFileName = Left(lpTempFileName, _
                         InStr(lpTempFileName, Chr(0)) - 1)
   Call Kill(lpTempFileName)
   If Len(UseExtension) > 0 Then
     lpTempFileName = Left(lpTempFileName, Len(lpTempFileName) - 3) & UseExtension
   End If
   GetUniqueFilename = lpTempFileName
 End Function

Private Sub AddDirSep(strPathName As String)
    If Right$(RTrim$(strPathName), 1) <> "\" Then
    strPathName = RTrim$(strPathName) & "\"
    End If
End Sub

Public Function LoadGif(sFile As String, aImg As Variant) As Long
On Error Resume Next
Dim lngFind As Long, lngPreviousFind As Long, strTempFile As String
   Dim hFile         As Long
   Dim sFileHeader   As String, strTemp As String
   Dim sBuff         As String
   Dim sPicsBuff     As String
   Dim TimeWait      As Long
   Dim bolDoLastImage As Boolean
TotalFrames = 0
If Dir$(sFile) = "" Or sFile = "" Then
    Exit Function
End If
MyPath = App.Path
AddDirSep MyPath
   If aImg.Count > 1 Then
      For i = 1 To aImg.Count - 1
         Unload aImg(i)
      Next i
   End If
   
  'load the gif into a string buffer
   hFile = FreeFile
    
   Open sFile For Binary Access Read As hFile
      sBuff = String(LOF(hFile), Chr(0))
      Get #hFile, , sBuff
   Close #hFile
        
  'find size of color table
    If Asc(Mid(sBuff, 11, 1)) And 128 Then
    lngFind = Asc(Mid(sBuff, 11, 1)) And 7
    lngFind = 3 * (2 ^ (lngFind + 1))
    End If
    lngFind = lngFind + 13
    sFileHeader = Left(sBuff, lngFind)
    
  'GIF?
    If Left$(sFileHeader, 3) <> "GIF" Then Exit Function
      
  'logical dimensions
  LogicalWidth = Asc(Mid(sBuff, 7, 1)) + Asc(Mid(sBuff, 8, 1)) * 256&
  LogicalHeight = Asc(Mid(sBuff, 9, 1)) + Asc(Mid(sBuff, 10, 1)) * 256&
      
  'temporary file
   hFile = FreeFile
   strTempFile = GetUniqueFilename(MyPath, "p" & Chr(0), "GIF")
   Open strTempFile For Binary As hFile
   
   'locate start of a frame
    lngFind = InStr(Len(sFileHeader) + 1, sBuff, sGifMagic) + 1
    
    'first image
If lngFind > 1 Then
    sPicsBuff = sFileHeader & Mid(sBuff, Len(sFileHeader) + 1, lngFind - (Len(sFileHeader) + 1)) & Trailer
    Put #hFile, 1, sPicsBuff
    Load aImg(1)
    aImg(1).Visible = True
    aImg(1).Tag = "10"
    aImg(1).Picture = LoadPicture(strTempFile)
    If aImg(1).Picture.Handle <> 0 Then
      TotalFrames = 1
    Else
      Unload aImg(1)
    End If
lngPreviousFind = lngFind
lngFind = InStr(lngPreviousFind + 1, sBuff, sGifMagic) + 1
Else
'only one image
lngPreviousFind = Len(sFileHeader) + 1
lngFind = Len(sBuff)
bolDoLastImage = True
End If

'search next image
Do While lngFind > 1
TotalFrames = TotalFrames + 1
Load aImg(TotalFrames)
aImg(TotalFrames).Visible = True
strTemp = Mid(sBuff, lngPreviousFind, lngFind - lngPreviousFind)
    sPicsBuff = sFileHeader & strTemp & Trailer
    Put #hFile, 1, sPicsBuff
    
'redraw?
aImg(TotalFrames).DrawStyle = (Asc(Mid(strTemp, 4, 1)) And 28) / 4

'load picture
aImg(TotalFrames).Picture = LoadPicture(strTempFile)
If bolDoLastImage Then
    If aImg(TotalFrames).Picture.Handle = 0 Then
      Unload aImg(TotalFrames)
      TotalFrames = TotalFrames - 1
      Exit Do
    End If
End If

'frame delay
TimeWait = ((Asc(Mid(strTemp, 5, 1))) + (Asc(Mid(strTemp, 6, 1)) * 256&)) * 10&
If TimeWait = 0 Then TimeWait = 1
If TimeWait > 65535 Then TimeWait = 65535
aImg(TotalFrames).Tag = CStr(TimeWait)

'position
If TotalFrames > 1 Then
aImg(TotalFrames).Left = aImg(1).Left + Asc(Mid(strTemp, 10, 1)) + (Asc(Mid(strTemp, 11, 1)) * 256&)
aImg(TotalFrames).Top = aImg(1).Top + Asc(Mid(strTemp, 12, 1)) + (Asc(Mid(strTemp, 13, 1)) * 256&)
End If
            
lngPreviousFind = lngFind
lngFind = InStr(lngPreviousFind + 1, sBuff, sGifMagic) + 1

'last image
If lngFind <= 1 Then
   If Not bolDoLastImage Then
      lngFind = Len(sBuff)
      bolDoLastImage = True
   End If
End If
Loop
    
   LoadGif = TotalFrames
   
   Close #hFile
   Kill strTempFile
   On Error GoTo 0
End Function




Attribute VB_Name = "TransparentBitblt"
Option Explicit
Private Const SRCCOPY = &HCC0020

'Sets the backcolour of a device context:
Private Declare Function SetBkColor Lib "gdi32" (ByVal hDC As Long, _
        ByVal crColor As Long) As Long
Public Function CreateMaskImage(ByRef picFrom As PictureBox, _
       ByRef picTo As PictureBox, Optional ByVal lTransparentColor _
       As Long = -1) As Boolean

   Dim lhDC As Long
   Dim lhBmp As Long
   Dim lhBmpOld As Long

   ' Create a monochrome DC & Bitmap of the
   ' same size as the source picture:
   lhDC = CreateCompatibleDC(0)
   If (lhDC <> 0) Then
      lhBmp = CreateCompatibleBitmap(lhDC, picFrom.ScaleWidth _
              , picFrom.ScaleHeight _
              )
      If (lhBmp <> 0) Then
         lhBmpOld = SelectObject(lhDC, lhBmp)
         ' Set the back 'colour' of the monochrome
         ' DC to the colour we wish to be transparent:
         If (lTransparentColor = -1) Then lTransparentColor = _
            picFrom.BackColor

         SetBkColor lhDC, lTransparentColor
         ' Copy from the from picture to the monochrome
         ' DC to create the mask:
         BitBlt lhDC, 0, 0, picFrom.ScaleWidth _
                , picFrom.ScaleHeight _
                , picFrom.hDC, 0, 0, SRCCOPY

         ' Now put the mask into picTo:
         BitBlt picTo.hDC, 0, 0, picFrom.ScaleWidth _
                , picFrom.ScaleHeight _
                , lhDC, 0, 0, SRCCOPY
         picTo.Refresh

         ' Clear up the bitmap we used to create
         ' the mask:
         SelectObject lhDC, lhBmpOld
         DeleteObject lhBmp
      End If
      ' Clear up the monochrome DC:
      DeleteObject lhDC
   End If
End Function


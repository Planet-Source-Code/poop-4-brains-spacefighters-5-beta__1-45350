Attribute VB_Name = "modCollision"

'**************************************
'Windows API/Global Declarations for :Ac
'     curate Collision Detection to Pixel leve
'     l
'**************************************
'loads (see below).
'**************************************
' Name: Accurate Collision Detection to
'     Pixel level
' Description:This code allows accurate
'     collision detection down to pixel level.


'     For use with Sprites, and using API call
    '     s to make this fast enough to work with
    '     a simple game.
' By: BigCalm
'
'
' Inputs:None
'
' Returns:None
'
'Assumes:None
'
'Side Effects:None
'This code is copyrighted and has limite
'     d warranties.
'Please see http://www.Planet-Source-Cod
'     e.com/xq/ASP/txtCodeId.21622/lngWId.1/qx
'     /vb/scripts/ShowCode.htm
'for details.
'**************************************



Type RECT
    Left As Long
    Top As Long
    Right As Long
    Bottom As Long
    End Type


Declare Function CreateCompatibleDC Lib "gdi32" (ByVal hdc As Long) As Long


Declare Function DeleteDC Lib "gdi32" (ByVal hdc As Long) As Long


Declare Function CreateCompatibleBitmap Lib "gdi32" (ByVal hdc As Long, ByVal nWidth As Long, ByVal nHeight As Long) As Long


Declare Function GetPixel Lib "gdi32" (ByVal hdc As Long, ByVal x As Long, ByVal y As Long) As Long


Declare Function BitBlt Lib "gdi32" (ByVal hDestDC As Long, ByVal x As Long, ByVal y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal hSrcDC As Long, ByVal xSrc As Long, ByVal ySrc As Long, ByVal dwRop As Long) As Long


Declare Function SelectObject Lib "gdi32" (ByVal hdc As Long, ByVal hObject As Long) As Long


Declare Function DeleteObject Lib "gdi32" (ByVal hObject As Long) As Long


Declare Function IntersectRect Lib "user32" (lpDestRect As RECT, lpSrc1Rect As RECT, lpSrc2Rect As RECT) As Long
    '---------------------------------------
    '     ----
    ' Collision Detection (Sprites)
    '---------------------------------------
    '     ----
    ' - Acknowledgement here goes to Richard
    '     Lowe (riklowe@hotmail.com) for his colli
    '     sion detection
    ' algorithm which I have used as the bas
    '     is of my collision detection algorithm.
    '     Most of the logic in
    ' here is radically different though, an
    '     d his algorithm originally didn't deallo
    '     cate memory properly ;-)
    ' - All X/Y/Width/Height values MUST be
    '     measured in pixels (ScaleMode = 3).
    ' - Compares bounding rectangles, and if
    '     they overlap, it goes to a pixel-by-pixe
    '     l comparison.
    'This therefore has detection down to th
    '     e pixel level.
    ' Function assumes you are using Masking
    '     sprites (not an unreasonable assumption,
    '     I'm sure you'll agree).
    ' - e.g. To test if collision has occurr
    '     ed between two sprites, one called "Ball
    '     ", the other "Bat":
    ' CollisionDetect(Ball.X,Ball.Y,Ball.Wid
    '     th, Ball.Height, 0, 0, Ball.MaskHdc, Bat
    '     .X, Bat.Y, Bat.Width, Bat.Height, 0, 0,
    '     Bat.MaskHdc)


Public Function CollisionDetect(ByVal x1 As Long, ByVal y1 As Long, ByVal X1Width As Long, ByVal Y1Height As Long, _
    ByVal Mask1LocX As Long, ByVal Mask1LocY As Long, ByVal Mask1Hdc As Long, ByVal x2 As Long, ByVal y2 As Long, _
    ByVal X2Width As Long, ByVal Y2Height As Long, ByVal Mask2LocX As Long, ByVal Mask2LocY As Long, _
    ByVal Mask2Hdc As Long) As Boolean
    ' I'm going to use RECT types to do this
    '     , so that the Windows API can do the har
    '     d bits for me.
    Dim MaskRect1 As RECT
    Dim MaskRect2 As RECT
    Dim DestRect As RECT
    Dim i As Long
    Dim j As Long
    Dim Collision As Boolean
    Dim MR1SrcX As Long
    Dim MR1SrcY As Long
    Dim MR2SrcX As Long
    Dim MR2SrcY As Long
    Dim hNewBMP As Long
    Dim hPrevBMP As Long
    Dim tmpObj As Long
    Dim hMemDC As Long
    MaskRect1.Left = x1
    MaskRect1.Top = y1
    MaskRect1.Right = x1 + X1Width
    MaskRect1.Bottom = y1 + Y1Height
    MaskRect2.Left = x2
    MaskRect2.Top = y2
    MaskRect2.Right = x2 + X2Width
    MaskRect2.Bottom = y2 + Y2Height
    i = IntersectRect(DestRect, MaskRect1, MaskRect2)


    If i = 0 Then
        CollisionDetect = False
    Else
        ' The two rectangles intersect, so let's
        '     go to a pixel by pixel comparison
        ' Set SourceX and Y values for both Mask
        '     HDC's...


        If x1 > x2 Then
            MR1SrcX = 0
            MR2SrcX = x1 - x2
        Else
            MR2SrcX = 0
            MR1SrcX = x2 - x1
        End If


        If y1 > y2 Then
            MR2SrcY = y1 - y2
            MR1SrcY = 0
        Else
            MR2SrcY = 0 ' here
            MR1SrcY = y2 - y1 - 1
        End If
        ' Allocate memory DC and Bitmap in which
        '     to do the comparison
        hMemDC = CreateCompatibleDC(Screen.ActiveForm.hdc)
        hNewBMP = CreateCompatibleBitmap(Screen.ActiveForm.hdc, DestRect.Right - DestRect.Left, DestRect.Bottom - DestRect.Top)
        hPrevBMP = SelectObject(hMemDC, hNewBMP)
        ' Blit the first sprite into it
        i = BitBlt(hMemDC, 0, 0, DestRect.Right - DestRect.Left, DestRect.Bottom - DestRect.Top, _
        Mask1Hdc, MR1SrcX + Mask1LocX, MR1SrcY + Mask1LocY, vbSrcCopy)
        ' Logical OR the second sprite with the
        '     first sprite
        i = BitBlt(hMemDC, 0, 0, DestRect.Right - DestRect.Left, DestRect.Bottom - DestRect.Top, _
        Mask2Hdc, MR2SrcX + Mask2LocX, MR2SrcY + Mask2LocY, vbSrcPaint)
        
        Collision = False


        For i = 0 To DestRect.Bottom - DestRect.Top - 1


            For j = 0 To DestRect.Right - DestRect.Left - 1


                If GetPixel(hMemDC, j, i) = 0 Then ' If there are any black pixels
                    Collision = True
                    Exit For
                End If
            Next


            If Collision = True Then
                Exit For
            End If
        Next
        CollisionDetect = Collision
        ' Destroy any allocated objects and DC's
        '
        tmpObj = SelectObject(hMemDC, hPrevBMP)
        tmpObj = DeleteObject(tmpObj)
        tmpObj = DeleteDC(hMemDC)
    End If
End Function

        



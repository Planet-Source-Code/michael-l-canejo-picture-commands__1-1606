<div align="center">

## Picture Commands


</div>

### Description

This includes the source code for Rotating a Picture 45 Degrees, Flipping a Picture Horizontally, and Flipping a Picture Vertically! Very usefull if your working on your own picture editing program. <br>Questions: mikecanejo@Hotmail.com
 
### More Info
 
'Please follow these simple steps. DONT SKIP ANYTHING!

'1.) Make a New Project in your 32bit Visual Basic

'2.) Add a New Module(Module1.Bas) to the New Project

'3.) Add a PictureBox to the Form(Picture1)and Insert a Picture in the 'PictureBox

'4.) Add a PictureBox to the Form(Picture2)and keep it blank

'5.) Add a CommandButton to the Form(Command1)

'6.) Add a CommandButton to the Form(Command2)

'7.) Add a CommandButton to the Form(Command3)

'Thats it, hope you followed EVERY step.


<span>             |<span>
---                |---
**Submitted On**   |
**By**             |[Michael L\. Canejo](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByAuthor/michael-l-canejo.md)
**Level**          |Advanced
**User Rating**    |4.0 (4 globes from 1 user)
**Compatibility**  |VB 4\.0 \(32\-bit\), VB 5\.0, VB 6\.0
**Category**       |[VB function enhancement](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByCategory/vb-function-enhancement__1-25.md)
**World**          |[Visual Basic](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByWorld/visual-basic.md)
**Archive File**   |[](https://github.com/Planet-Source-Code/michael-l-canejo-picture-commands__1-1606/archive/master.zip)

### API Declarations

```
'Copy and Paste the following into the New Module(Module1.Bas)..not in the form!
Global Const SRCCOPY = &HCC0020
Global Const Pi = 3.14159265359
Declare Function SetPixel Lib "gdi32" (ByVal hdc As Long, ByVal x As Long, ByVal y As Long, ByVal crColor As Long) As Long
Declare Function GetPixel Lib "gdi32" (ByVal hdc As Long, ByVal x As Long, ByVal y As Long) As Long
Declare Function StretchBlt Lib "gdi32" (ByVal hdc As Long, ByVal x As Long, ByVal y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal hSrcDC As Long, ByVal xSrc As Long, ByVal ySrc As Long, ByVal nSrcWidth As Long, ByVal nSrcHeight As Long, ByVal dwRop As Long) As Long
Sub bmp_rotate(pic1 As PictureBox, pic2 As PictureBox, ByVal theta!)
 'Rotate the image in a picture box.
 'pic1 is the picture box with the bitmap to rotate
 'pic2 is the picture box to receive the rotated bitmap
 'theta is the angle of rotation
 Dim c1x As Integer, c1y As Integer
 Dim c2x As Integer, c2y As Integer
 Dim a As Single
 Dim p1x As Integer, p1y As Integer
 Dim p2x As Integer, p2y As Integer
 Dim n As Integer, r As Integer
 c1x = pic1.ScaleWidth \ 2
 c1y = pic1.ScaleHeight \ 2
 c2x = pic2.ScaleWidth \ 2
 c2y = pic2.ScaleHeight \ 2
 If c2x < c2y Then n = c2y Else n = c2x
 n = n - 1
 pic1hDC% = pic1.hdc
 pic2hDC% = pic2.hdc
 For p2x = 0 To n
 For p2y = 0 To n
 If p2x = 0 Then a = Pi / 2 Else a = Atn(p2y / p2x)
 r = Sqr(1& * p2x * p2x + 1& * p2y * p2y)
 p1x = r * Cos(a + theta!)
 p1y = r * Sin(a + theta!)
 c0& = GetPixel(pic1hDC%, c1x + p1x, c1y + p1y)
 c1& = GetPixel(pic1hDC%, c1x - p1x, c1y - p1y)
 c2& = GetPixel(pic1hDC%, c1x + p1y, c1y - p1x)
 c3& = GetPixel(pic1hDC%, c1x - p1y, c1y + p1x)
 If c0& <> -1 Then xret& = SetPixel(pic2hDC%, c2x + p2x, c2y + p2y, c0&)
 If c1& <> -1 Then xret& = SetPixel(pic2hDC%, c2x - p2x, c2y - p2y, c1&)
 If c2& <> -1 Then xret& = SetPixel(pic2hDC%, c2x + p2y, c2y - p2x, c2&)
 If c3& <> -1 Then xret& = SetPixel(pic2hDC%, c2x - p2y, c2y + p2x, c3&)
 Next
 t% = DoEvents()
 Next
End Sub
```


### Source Code

```
'Ok, now just Copy and Paste everything here into the Form1..not in the Bas!
'<Start Copying>
Private Sub Command1_Click()
 'flip horizontal
 Picture2.Cls
 px% = Picture1.ScaleWidth
 py% = Picture1.ScaleHeight
 retval% = StretchBlt(Picture2.hdc, px%, 0, -px%, py%, Picture1.hdc, 0, 0, px%, py%, SRCCOPY)
End Sub
Private Sub Command2_Click()
 'flip vertical
 Picture2.Cls
 px% = Picture1.ScaleWidth
 py% = Picture1.ScaleHeight
 retval% = StretchBlt(Picture2.hdc, 0, py%, px%, -py%, Picture1.hdc, 0, 0, px%, py%, SRCCOPY)
End Sub
Private Sub Command3_Click()
 'rotate 45 degrees
 Picture2.Cls
 Call bmp_rotate(Picture1, Picture2, 3.14 / 4)
End Sub
Private Sub Form_Load()
Command1.Caption = "Flip Horizontal"
Command2.Caption = "Flip Vertical"
Command3.Caption = "Rotate 45 Degrees"
 Picture1.ScaleMode = 3
 Picture2.ScaleMode = 3
End Sub
'<Stop Copying...END>
```


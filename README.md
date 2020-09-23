<div align="center">

## Load bitmaps straight into memory


</div>

### Description

Do you have a game with too many pictureboxes for your graphics. Want to make your game really proffessional by loading them straight into RAM. Even returns all the properties of your bitmap. Can be used with 3 liones of code. This is to be used with bitblt. Not compatible with Windows NT due to the loadfromfile constant being restricted or something.
 
### More Info
 
Use it like this:

Private sub sub Form_Load()

Me.show

Dim lngDC as long

Dim bmpProperties as BITMAP

lngDC = GenerateDC(C:\Bitmap.Bmp,bmpProperties)

Msgbox BmpProperties.bmwidth

me.scalemode = vbpixels

Bitblt me.hdc,0,0,me.scalewidth,me.scaleheight,lngDC,0,0,vbsrccopy

deletegenerateddc lngdc

end sub

You place the code in a module.

Puts a bitmap in memory and returns its properties and a DC to it. Then when you are finished just delete the dc using the provided sub.

Remember to delete the DC or you might start losing some resources. Dont attempt to use on Windows NT.


<span>             |<span>
---                |---
**Submitted On**   |
**By**             |[IRBMe](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByAuthor/irbme.md)
**Level**          |Intermediate
**User Rating**    |4.8 (19 globes from 4 users)
**Compatibility**  |VB 6\.0
**Category**       |[Games](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByCategory/games__1-38.md)
**World**          |[Visual Basic](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByWorld/visual-basic.md)
**Archive File**   |[](https://github.com/Planet-Source-Code/irbme-load-bitmaps-straight-into-memory__1-28270/archive/master.zip)

### API Declarations

Various. See code


### Source Code

```
Option Explicit
Private Declare Function CreateCompatibleDC Lib "gdi32" (ByVal hdc As Long) As Long
Private Declare Function DeleteDC Lib "gdi32" (ByVal hdc As Long) As Long
Private Declare Function DeleteObject Lib "gdi32" (ByVal hObject As Long) As Long
Private Declare Function GetObjectAPI Lib "gdi32" Alias "GetObjectA" (ByVal hObject As Long, ByVal nCount As Long, lpObject As Any) As Long
Private Declare Function LoadImage Lib "user32" Alias "LoadImageA" (ByVal hInst As Long, ByVal lpsz As String, ByVal un1 As Long, ByVal n1 As Long, ByVal n2 As Long, ByVal un2 As Long) As Long
Private Declare Function SelectObject Lib "gdi32" (ByVal hdc As Long, ByVal hObject As Long) As Long
Private Const IMAGE_BITMAP As Long = 0
Private Const LR_LOADFROMFILE As Long = &H10
Private Const LR_CREATEDIBSECTION As Long = &H2000
Public Type BITMAP
  bmType As Long
  bmWidth As Long
  bmHeight As Long
  bmWidthBytes As Long
  bmPlanes As Integer
  bmBitsPixel As Integer
  bmBits As Long
End Type
Public Function GenerateDC(ByVal FileName As String, BitmapProperties As BITMAP) As Long
Dim DC As Long
Dim hBitmap As Long
DC = CreateCompatibleDC(0)
If DC < 1 Then
  GenerateDC = 0
  Exit Function
End If
hBitmap = LoadImage(0, FileName, IMAGE_BITMAP, 0, 0, LR_LOADFROMFILE Or LR_CREATEDIBSECTION)
If hBitmap = 0 Then
  DeleteDC DC
  GenerateDC = 0
  Exit Function
End If
GetObjectAPI hBitmap, Len(BitmapProperties), BitmapProperties
SelectObject DC, hBitmap
GenerateDC = DC
DeleteObject hBitmap
End Function
Public Function DeleteGeneratedDC(DC As Long) As Long
If DC > 0 Then
  DeleteGeneratedDC = DeleteDC(DC)
Else
  DeleteGeneratedDC = 0
End If
End Function
'Gimme somefeedback and votes please. Thats the only time I'm gonna ask
```


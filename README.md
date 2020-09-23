<div align="center">

## Capturing the Screen


</div>

### Description

Capture a screen in a window, this one actually works...
 
### More Info
 
'Inputs: None

You need to create a form, add 2 menu items, item1 and item2, then add 2 picture boxs one named piccover and one named picfinal, make them the same size, and lay them right on top of each other... the size of them is the size of the screen captured... so if you want a 1024x768 screen captured be sure to size the picture boxes as big as you can.. The two menu items you can call whatever you like.. Capture screen for the first, and Clear window for the second if you like...


<span>             |<span>
---                |---
**Submitted On**   |
**By**             |[TK](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByAuthor/tk.md)
**Level**          |Unknown
**User Rating**    |1.0 (1 globe from 1 user)
**Compatibility**  |VB 5\.0, VB 6\.0
**Category**       |[Custom Controls/ Forms/  Menus](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByCategory/custom-controls-forms-menus__1-4.md)
**World**          |[Visual Basic](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByWorld/visual-basic.md)
**Archive File**   |[](https://github.com/Planet-Source-Code/tk-capturing-the-screen__1-948/archive/master.zip)

### API Declarations

```
'Make these one line, this window isn't big enough
Private Declare Function GetDC% Lib "USER32" (ByVal HWnd%)
Private Declare Function ReleaseDC% Lib "USER32" (ByVal HWnd%, ByVal HDC%)
Private Declare Function GetDesktopWindow% Lib "USER32" ()
Private Declare Function BitBlt% Lib "GDI32" (ByVal DestDC%, ByVal X%, ByVal Y%, ByVal W%, ByVal H%, ByVal SrcHDC%, ByVal SrcX%, ByVal SrcY%, ByVal Rop&)
Private Declare Function DeleteDC Lib "GDI32" (ByVal HDC As Long) As Long
Const SRCCOPY = &HCC0020
```


### Source Code

```
Private Sub GrabScreen()
'I wont format this because this box doesn't allow tabbing, my apologies...
PicFinal.Cls
DeleteDC (HwndSrc%)
HwndSrc% = GetDesktopWindow()
HSrcDC% = GetDC(HwndSrc%)
'BitBlt requires coordinates in pixels.
HDestDC% = PicFinal.HDC
DWRop& = SRCCOPY
Suc% = BitBlt(HDestDC%, 0, 0, 1024, 768, HSrcDC%, 0, 0, DWRop&)
Dmy% = ReleaseDC(HwndSrc%, HSrcDC%)
PicCover.Picture = PicFinal.Image
DeleteDC (HwndSrc%)
End Sub
Private Sub Item2_Click()
Capture.Hide
Capture.Visible = False
GrabScreen
Capture.Visible = True
End Sub
Private Sub Item3_Click()
Cls
PicFinal.Cls
PicCover.Cls
PicFinal.Refresh
PicCover.Refresh
DeleteDC (HwndSrc%)
End Sub
```


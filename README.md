<div align="center">

## Image Thumbnail Viewer


</div>

### Description

This is a very simple way to view a thumbnail of an image file while maintaining the aspect ration.
 
### More Info
 
'String - The full path to the image file

'PictureBox1 - Temporary holding place for the image

'PictureBox2 - Final resting place for the image


<span>             |<span>
---                |---
**Submitted On**   |
**By**             |[Ben White](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByAuthor/ben-white.md)
**Level**          |Beginner
**User Rating**    |4.7 (14 globes from 3 users)
**Compatibility**  |VB 6\.0
**Category**       |[Graphics](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByCategory/graphics__1-46.md)
**World**          |[Visual Basic](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByWorld/visual-basic.md)
**Archive File**   |[](https://github.com/Planet-Source-Code/ben-white-image-thumbnail-viewer__1-33677/archive/master.zip)





### Source Code

```
Option Explicit
Public Sub ViewImage(strFile As String, picTemp As PictureBox, picTarget As PictureBox)
  Dim x&, y&, x1&, y1&, z1!
  Dim sNoPreview$
  On Error GoTo ErrorHandler
 'Set default stuffs
 picTarget.Cls
 picTarget.AutoRedraw = True
 picTemp.Visible = False
 picTemp.AutoSize = True
 'get target sizing info
 x = picTarget.width
 y = picTarget.height
 'Load the image
 picTemp.Picture = LoadPicture(strFile)
 'get source sizing info
 x1 = picTemp.width
 y1 = picTemp.height
 'Determin conversion ratio to use
 z1 = IIf(x / x1 * y1 < y, x / x1, y / y1)
 'Calculate new image size
 x1 = x1 * z1
 y1 = y1 * z1
 'Draw Image
 picTarget.PaintPicture picTemp.Picture, (x - x1) / 2, (y - y1) / 2, x1, y1
 Exit Sub
ErrorHandler:
 'set temp image to nothing
 picTemp.Picture = LoadPicture()
 'Display default error message
 sNoPreview = "No Preview Available"
 picTarget.CurrentX = x / 2 - picTarget.TextWidth(sNoPreview) / 2
 picTarget.CurrentY = y / 2 - picTarget.TextHeight(sNoPreview) / 2
 picTarget.Print sNoPreview
End Sub
```


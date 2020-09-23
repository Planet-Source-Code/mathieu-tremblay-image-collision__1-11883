<div align="center">

## Image Collision


</div>

### Description

Verify if two images will hit if you do move one of the image of X left and X right. Very useful when programming actions, sports or rpg games.
 
### More Info
 
MovingImage = the image to move

moveLeft  = the Left movement of the MovingImage

moveTop   = the Top movement of the MovingImage

StaticImage = the image you don't want to hit

Return if the two images would hit ( True or False ).

The MovingImage will automaticly move if you don't erase the two folling lines of code:

MovingImage.Left = MovingLeft

MovingImage.Top = MovingTop


<span>             |<span>
---                |---
**Submitted On**   |
**By**             |[Mathieu Tremblay](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByAuthor/mathieu-tremblay.md)
**Level**          |Intermediate
**User Rating**    |4.3 (17 globes from 4 users)
**Compatibility**  |VB 5\.0, VB 6\.0
**Category**       |[Miscellaneous](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByCategory/miscellaneous__1-1.md)
**World**          |[Visual Basic](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByWorld/visual-basic.md)
**Archive File**   |[](https://github.com/Planet-Source-Code/mathieu-tremblay-image-collision__1-11883/archive/master.zip)





### Source Code

```
Public Function CollisionMovingImage(MovingImage As Variant, moveLeft As Integer, moveTop As Integer, Optional StaticImage As Variant) As Boolean
On Error GoTo ErrHandler:
'If one of the parameters is not found or
'some error happen in the function, it will
'then exit.
  Dim MovingLeft, MovingRight, MovingTop, MovingBottom As Integer
  'The Moving variables are used to get infos about the
  'MovingImage.
  MovingLeft = MovingImage.Left + moveLeft
  MovingRight = (MovingImage.Left + moveLeft) + MovingImage.Width
  MovingTop = MovingImage.Top + moveTop
  MovingBottom = (MovingImage.Top + moveTop) + MovingImage.Height
  Dim okLeft, okTop As Boolean
  ' okLeft is use to see if the MovingImage has a point
  ' of its width in commun with the StaticImage. The
  ' okTop is used to see if it happens with the height.
  okLeft = True
  okTop = True
  'They are set to true by default to allow the moving
  'of the MovingImage if there is no StaticImage.
  If IsMissing(StaticImage) = False Then
  'Execute the verification only if the
  'StaticImage argument is used.
    Dim StaticLeft, StaticRight, StaticTop, StaticBottom As String
    'The Static variables are used to get infos about
    'the StaticImage.
    StaticLeft = StaticImage.Left
    StaticRight = StaticImage.Left + StaticImage.Width
    StaticTop = StaticImage.Top
    StaticBottom = StaticImage.Top + StaticImage.Height
    Dim i As Integer
    'Verify if the MovingImage has a point
    'of its width in commun with the StaticImage.
    For i = StaticLeft To StaticRight
      If (MovingLeft = i) Or (MovingRight = i) Then
        okLeft = False
      End If
    Next i
    'Verify if the MovingImage has a point of
    'its height in commun with the StaticImage.
    For i = StaticTop To StaticBottom
      If (MovingBottom = i) Or (MovingTop = i) Then
        okTop = False
      End If
    Next i
    'Don't move the MovingPicture if there
    'would be a collision.
    If okTop = False And okLeft = False Then
      'Return true because the two objects
      'would have a commun point.
      CollisionMovingImage = True
      GoTo ErrHandler:
    End If
  End If
  'Move the MovingImage...
  'You could remove the two following lines if you
  'wanted the function to only tell you if there would
  'be a collision or no.
  MovingImage.Left = MovingLeft
  MovingImage.Top = MovingTop
  'Return false because there have been no collision
  CollisionMovingImage = False
ErrHandler:
End Function
```


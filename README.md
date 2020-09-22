<div align="center">

## Auto resize that works


</div>

### Description

I needed a *good* auto resizer to minimize the time spent on resize code. I tried a few auto resizers, but none of them worked the way I wanted. So, I threw together this little piece of code. Unlike other resizers making assumptions on how to resize your controls, this code makes no assumptions. You the programmer are in total control of the resize behavior of each control on the form. IMPORTANT NOTES BELOW!
 
### More Info
 
IMPORTANT!! This code needs the tag property of your controls. If your tag properties are in use, then you'll have to devise with another method. For each control, set its tag property to the code for its resize behavior.

Codes:

sX = Stretch on X axis

sY = Stretch on Y axis

sXY = Stretch on Both axis

rY = Move relative to Y axis

rX = Move relative to X axis

rXY = Move relative to XY axis

More notes in the code!

Works for most resizing... may not function properly with OLE objects, but its not a big deal.


<span>             |<span>
---                |---
**Submitted On**   |
**By**             |[Chris Bradford](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByAuthor/chris-bradford.md)
**Level**          |Intermediate
**User Rating**    |4.8 (24 globes from 5 users)
**Compatibility**  |VB 5\.0, VB 6\.0
**Category**       |[Miscellaneous](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByCategory/miscellaneous__1-1.md)
**World**          |[Visual Basic](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByWorld/visual-basic.md)
**Archive File**   |[](https://github.com/Planet-Source-Code/chris-bradford-auto-resize-that-works__1-24503/archive/master.zip)

### API Declarations

```
Public Type CtlAdj
  adjX As Long
  adjY As Long
End Type
```


### Source Code

```
'*** put the following code in a module or something normal
Public Sub ResizeOMatic(frm As Form, adj() As CtlAdj)
  '** ResizeOMatic :: this sub moves and resizes controls on the form based
  '** on the adjustment data passed. Each element of the adj array should be
  '** in sequence as long as VB enumerates the controls in the same order as it
  '** did when the adj array was built (sub RegisterForm)
  Dim tmpControl As Control
  Dim index As Long
  On Error Resume Next        'keepin it real
  index = 0
  For Each tmpControl In frm
    index = index + 1
    Select Case LCase$(tmpControl.Tag)
      Case "rx"      'relative X
        tmpControl.Left = frm.width - tmpControl.width - adj(index).adjX
      Case "ry"      'relative Y
        tmpControl.Top = frm.height - tmpControl.height - adj(index).adjY
      Case "rxy"     'relative XY
        tmpControl.Left = frm.width - tmpControl.width - adj(index).adjX
        tmpControl.Top = frm.height - tmpControl.height - adj(index).adjY
      Case "sx"      'stretch X
        tmpControl.width = frm.width - tmpControl.Left - adj(index).adjX
      Case "sy"      'stretch Y
        tmpControl.height = frm.height - tmpControl.Top - adj(index).adjY
      Case "sxy"     'stretch XY
        tmpControl.width = frm.width - tmpControl.Left - adj(index).adjX
        tmpControl.height = frm.height - tmpControl.Top - adj(index).adjY
    End Select
  Next
End Sub
Public Sub RegisterForm(frm As Form, width As Long, height As Long, adj() As CtlAdj)
  '** RegisterForm :: this sub enumerates the controls on the form and records
  '** the positions of the bottom right corner of the control. We have to pass the
  '** width and height parameters (initial point of reference) because MDI
  '** automagically sizes forms. The adjustment data is used in Sub ResizeOMatic
  Dim tmpControl As Control
  ReDim adj(0)
  On Error Resume Next                 'keepin it real
  For Each tmpControl In frm
    ReDim Preserve adj(UBound(adj) + 1)
    adj(UBound(adj)).adjX = width - (tmpControl.Left + tmpControl.width)
    adj(UBound(adj)).adjY = height - (tmpControl.Top + tmpControl.height)
  Next
End Sub
'*********** The following code is a form
'*********** demonstrating how to use it
Private Sizedata() As CtlAdj
Private Sub Form_Load()
  '** load your stuff here
  'call this near the end of the form_load()
'Note: On MDI child forms, you should manually
'specify the width and height to your design time
'size to keep proper proportions
  RegisterForm Me, Me.Width, Me.Height, Sizedata()
End Sub
Private Sub Form_Resize()
  ResizeOMatic Me, Sizedata()
End Sub
```


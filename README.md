<div align="center">

## Acces control thru variable


</div>

### Description

Be able to access multiple controls which require the same task with minimal code. Basically I wanted to be able to access controls without having to write numerous bits and peices to reach multiple controls with multiple forms and could not find what I needed so I decided to post. Here you are creating 2 control arrays with labels and all the code will demonstrate is randomly changeing the color of a randomly selected label.

This is my first post so if a bit lame hope you at least find it usefull.
 
### More Info
 
You need to create a form with the following items.

Timer1 Set Interval = 1000

Label1 (0 - 7) Control Array

Label2 (0 - 7) Control Array


<span>             |<span>
---                |---
**Submitted On**   |
**By**             |[James Travis](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByAuthor/james-travis.md)
**Level**          |Intermediate
**User Rating**    |4.3 (13 globes from 3 users)
**Compatibility**  |VB 5\.0, VB 6\.0
**Category**       |[Miscellaneous](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByCategory/miscellaneous__1-1.md)
**World**          |[Visual Basic](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByWorld/visual-basic.md)
**Archive File**   |[](https://github.com/Planet-Source-Code/james-travis-acces-control-thru-variable__1-13300/archive/master.zip)





### Source Code

```
Private Sub Timer1_Timer()
  Dim CtrlName As String, CtrlIdx As Integer
  Dim ClrSlct As Integer, ClrSlcted As Long
  Dim LblRow As Integer
  Randomize
  CtrlIdx = Rnd * 7
  Randomize
  ClrSlct = Rnd * 3
  Select Case ClrSlct
    Case 0
      ClrSlcted = &HFF00&   'Green
    Case 1
      ClrSlcted = &HFF&    'Red
    Case 2
      ClrSlcted = &HFF0000  'Blue
    Case 3
      ClrSlcted = &HFFFF&   'Yellow
  End Select
  Randomize
  LblRow = (Rnd * 1) + 1
  CtrlName = "Label" & LblRow
  Form1.Controls(CtrlName).Item(CtrlIdx).BackColor = ClrSlcted
End Sub
```


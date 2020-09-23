<div align="center">

## Print Layout


</div>

### Description

This piece of code loops through the controls on a form and sends the contents to the default printer in a layout similar to the screen. Therefore your forms contents are printed in the same positions as they are on screen. I use this to print out a simple record report on the db application I am working on. Works with labels, text boxes and list boxes etc.
 
### More Info
 


<span>             |<span>
---                |---
**Submitted On**   |
**By**             |[ronniec23](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByAuthor/ronniec23.md)
**Level**          |Intermediate
**User Rating**    |5.0 (10 globes from 2 users)
**Compatibility**  |VB 6\.0
**Category**       |[Miscellaneous](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByCategory/miscellaneous__1-1.md)
**World**          |[Visual Basic](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByWorld/visual-basic.md)
**Archive File**   |[](https://github.com/Planet-Source-Code/ronniec23-print-layout__1-45597/archive/master.zip)





### Source Code

```
Sub Print_rec()
On Error GoTo print_err
With Printer
  .Orientation = 1 '1 = portrait, 2 = landscape
  .CurrentX = 3000 ' move text over and down for title
  .CurrentY = 1440
  .FontBold = True
  .FontSize = 12
  Printer.Print "Record Details"
  .FontSize = 9 'Change back font size
  .FontBold = False 'Change back to none bold font
'This section loops through the controls on the screen and prints the contents of the control.
' I used the tag property to filter controls as there was some controls on the screen I didnt want printing (buttons, check boxes etc.)
For Each Control In Me.Controls
  If Control.Tag = "prt" Then
  .CurrentX = Control.Left + 250 ' sets the position for printing (+ 250 move's it in about 1cm from side of sheet)
  .CurrentY = Control.Top + 2400 ' + 2400 allows space for title
  If Control.Name Like "lbl*" Then
    Printer.Print Control.Caption & ":"  'print label captions and a ":"
  Else
  Printer.Print Control.Text ' prints contents of text box
  End If
  End If
Next Control
.EndDoc
End With
MsgBox "Printed!"
Exit Sub
print_err:
  MsgBox "Error in printing tender details."
  Exit Sub
End Sub
```


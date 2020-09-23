<div align="center">

## WordWrap in only 5 codelines


</div>

### Description

I hope that this is the shortest and easiest wordwrap-function in vb you have ever seen, that you enjoy it and use it in all your projects :-)
 
### More Info
 


<span>             |<span>
---                |---
**Submitted On**   |
**By**             |[Max Christian Pohle](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByAuthor/max-christian-pohle.md)
**Level**          |Advanced
**User Rating**    |4.8 (58 globes from 12 users)
**Compatibility**  |VB 5\.0, VB 6\.0
**Category**       |[String Manipulation](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByCategory/string-manipulation__1-5.md)
**World**          |[Visual Basic](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByWorld/visual-basic.md)
**Archive File**   |[](https://github.com/Planet-Source-Code/max-christian-pohle-wordwrap-in-only-5-codelines__1-62266/archive/master.zip)





### Source Code

```
Option Explicit
Private Sub Form_Load()
  MsgBox WordWrap("This is a long testtext that doesn't make any sense really. But I hope you will enjoy my example and I do not know what I can write any more. This must be enough", 20), vbOKOnly + vbInformation, "WordWrap"
End Sub
Function WordWrap(ByVal Text As String, Optional ByVal MaxLineLen As Integer = 70)
  Dim i As Integer
  For i = 1 To Len(Text) / MaxLineLen
    Text = Mid(Text, 1, MaxLineLen * i - 1) & Replace(Text, " ", vbCrLf, MaxLineLen * i, 1, vbTextCompare)
  Next i
  WordWrap = Text
End Function
```


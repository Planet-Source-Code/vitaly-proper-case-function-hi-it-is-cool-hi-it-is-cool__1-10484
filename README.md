<div align="center">

## Proper case function\! hi\! iT is CoOl \-\> Hi\! It is cool


</div>

### Description

To fix the input of the user. For example if the user inputs a bad case sentance it will make it normal:

what do yOu think? -> What do you think?

It checks for ./?/!. Have fun.
 
### More Info
 
Bad case string

Proper case string


<span>             |<span>
---                |---
**Submitted On**   |
**By**             |[Vitaly](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByAuthor/vitaly.md)
**Level**          |Beginner
**User Rating**    |3.6 (18 globes from 5 users)
**Compatibility**  |VB 3\.0, VB 4\.0 \(16\-bit\), VB 4\.0 \(32\-bit\), VB 5\.0, VB 6\.0, VB Script, ASP \(Active Server Pages\) 
**Category**       |[Coding Standards](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByCategory/coding-standards__1-43.md)
**World**          |[Visual Basic](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByWorld/visual-basic.md)
**Archive File**   |[](https://github.com/Planet-Source-Code/vitaly-proper-case-function-hi-it-is-cool-hi-it-is-cool__1-10484/archive/master.zip)





### Source Code

```
Function Proper(Text As String)
Dim FirstLetter%, I%, BadFormat%, J%, Loc%, Sta%
Dim BreakL$, BreakR$
'------------------------
Text = LCase(Text)
FirstLetter = Asc(Mid(Text, 1, 1))
If FirstLetter >= 97 And FirstLetter <= 122 Then
  Text = Right(Text, Len(Text) - 1)
  Text = Chr(FirstLetter - 32) & Text
End If
For I = 1 To Len(Text) - 2
  If Mid(Text, I, 1) = "." And Asc(Mid(Text, I + 2, 1)) >= 97 And Asc(Mid(Text, I + 2, 1)) <= 122 _
   Then BadFormat = BadFormat + 1
  If Mid(Text, I, 1) = "!" And Asc(Mid(Text, I + 2, 1)) >= 97 And Asc(Mid(Text, I + 2, 1)) <= 122 _
   Then BadFormat = BadFormat + 1
  If Mid(Text, I, 1) = "?" And Asc(Mid(Text, I + 2, 1)) >= 97 And Asc(Mid(Text, I + 2, 1)) <= 122 _
   Then BadFormat = BadFormat + 1
Next I
Loc = 1
For J = 1 To BadFormat
  Sta = 200
  If InStr(Loc, Text, ".") <> 0 Then Sta = InStr(Loc, Text, ".")
  If InStr(Loc, Text, "?") < Sta And InStr(Loc, Text, "?") <> 0 _
   Then Sta = InStr(Loc, Text, "?")
  If InStr(Loc, Text, "!") < Sta And InStr(Loc, Text, "!") <> 0 _
   Then Sta = InStr(Loc, Text, "!")
  For I = Sta To Len(Text)
    If Asc(Mid(Text, I, 1)) >= 97 And Asc(Mid(Text, I, 1)) <= 122 Then
      Loc = I + 1
      FirstLetter = Asc(Mid(Text, I, 1))
      BreakL = Left(Text, I - 1)
      BreakR = Right(Text, Len(Text) - I)
      Text = BreakL & Chr(FirstLetter - 32) & BreakR
      Exit For
    End If
  Next I
Next J
Proper = Text
End Function
```


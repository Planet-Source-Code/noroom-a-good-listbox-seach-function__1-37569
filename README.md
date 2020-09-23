<div align="center">

## A Good ListBox Seach function


</div>

### Description

This code searches a listbox for a given string, and returns the the ListIndex of the Item that matches the given string.
 
### More Info
 
The string you want to find in the listbox

if found: the ListIndex that matches the given string

If not found: returns -1


<span>             |<span>
---                |---
**Submitted On**   |
**By**             |[noroom](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByAuthor/noroom.md)
**Level**          |Beginner
**User Rating**    |3.3 (10 globes from 3 users)
**Compatibility**  |VB 5\.0, VB 6\.0
**Category**       |[Miscellaneous](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByCategory/miscellaneous__1-1.md)
**World**          |[Visual Basic](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByWorld/visual-basic.md)
**Archive File**   |[](https://github.com/Planet-Source-Code/noroom-a-good-listbox-seach-function__1-37569/archive/master.zip)





### Source Code

```
Private Function SearchList(ToSearch As String, lstList As ListBox) As Integer
Dim i As Integer
SearchList = -1
For i = 0 To lstList.ListCount - 1
 If LCase(lstList.List(i)) = LCase(ToSearch) Then
  SearchList = i
  Exit For
 End If
Next i
End Function
```


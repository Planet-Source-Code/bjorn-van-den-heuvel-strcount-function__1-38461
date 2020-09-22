<div align="center">

## StrCount function


</div>

### Description

Function that counts the occurrance of a given phrase in a larger string.

(I needed this function and couldn't find anything like it in MSDN for VB)
 
### More Info
 
cToSearch: The string to search.

cSearchPhrase: The phrase to count

Included the check on the length of the SearchPhrase, in order to prevent a division by zero.

The number of occurrances found.


<span>             |<span>
---                |---
**Submitted On**   |
**By**             |[Bjorn van den Heuvel](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByAuthor/bjorn-van-den-heuvel.md)
**Level**          |Advanced
**User Rating**    |4.4 (31 globes from 7 users)
**Compatibility**  |VB 3\.0, VB 4\.0 \(16\-bit\), VB 4\.0 \(32\-bit\), VB 5\.0, VB 6\.0, VB Script, ASP \(Active Server Pages\) , VBA MS Access, VBA MS Excel
**Category**       |[String Manipulation](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByCategory/string-manipulation__1-5.md)
**World**          |[Visual Basic](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByWorld/visual-basic.md)
**Archive File**   |[](https://github.com/Planet-Source-Code/bjorn-van-den-heuvel-strcount-function__1-38461/archive/master.zip)





### Source Code

```
Public Function StrCount(ByVal cToSearch As String, ByVal cSearchPhrase As String) As Long
 Dim nDifference As Long  ' The difference in length after the replace
 ' Is there anything to search?
 If Len(cSearchPhrase) > 0 Then
  nDifference = Len(cToSearch) - Len(Replace(cToSearch, cSearchPhrase, ""))
  StrCount = nDifference / Len(cSearchPhrase)
 Else
  StrCount = 0
 End If
End Function
```


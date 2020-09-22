<div align="center">

## Clear Documents Menu


</div>

### Description

Clears the documents menu in Windows 95/98/NT.
 
### More Info
 


<span>             |<span>
---                |---
**Submitted On**   |
**By**             |[Adam Petrie](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByAuthor/adam-petrie.md)
**Level**          |Unknown
**User Rating**    |5.0 (10 globes from 2 users)
**Compatibility**  |VB 4\.0 \(32\-bit\), VB 5\.0, VB 6\.0
**Category**       |[Miscellaneous](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByCategory/miscellaneous__1-1.md)
**World**          |[Visual Basic](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByWorld/visual-basic.md)
**Archive File**   |[](https://github.com/Planet-Source-Code/adam-petrie-clear-documents-menu__1-2657/archive/master.zip)





### Source Code

```
Option Explicit
Private Declare Function SHAddToRecentDocs Lib "Shell32" (ByVal lFlags As Long, ByVal lPv As Long) As Long
Private Sub Command1_Click()
SHAddToRecentDocs 0, 0 ' Clear All Items Under The Documents Menu
End Sub
```


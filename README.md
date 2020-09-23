<div align="center">

## \*\*\*Make ANY Folder The Root Of Windows Explorer\*\*\*


</div>

### Description

My Code calls The Windows Explorer with the switch "e,/root," and makes any folder you want the root of the windows explorer
 
### More Info
 
It's a little slow calling the explorer cause i shell it, if you know better please edit at will


<span>             |<span>
---                |---
**Submitted On**   |
**By**             |[Jim](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByAuthor/jim.md)
**Level**          |Unknown
**User Rating**    |4.2 (21 globes from 5 users)
**Compatibility**  |VB 4\.0 \(16\-bit\), VB 4\.0 \(32\-bit\), VB 5\.0, VB 6\.0
**Category**       |[Windows API Call/ Explanation](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByCategory/windows-api-call-explanation__1-39.md)
**World**          |[Visual Basic](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByWorld/visual-basic.md)
**Archive File**   |[](https://github.com/Planet-Source-Code/jim-make-any-folder-the-root-of-windows-explorer__1-3096/archive/master.zip)





### Source Code

```
'it's a module
'i went a little DIM crazy with the
'variables but it's still good code...enjoy
Public Sub eRoot(rootpath As String, fldrs As Boolean)
'fldrs is the folders switch, monkey with it and see what you get
On Error Resume Next
Dim EX, ARGU, path, X
If fldrs = True Then
EX = "explorer.exe"
ARGU = " /e,/root, "
path = rootpath$
X = Shell(EX & ARGU & path, 1)
ElseIf fldrs = False Then
EX = "explorer.exe"
ARGU = " n/e,/,root, "
path = rootpath$
X = Shell(EX & ARGU & path, 1)
End If
End Sub
```


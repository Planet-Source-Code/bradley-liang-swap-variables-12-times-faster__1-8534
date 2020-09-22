<div align="center">

## Swap Variables 12 times Faster\!


</div>

### Description

Swap Variable1 for Variable2 using API, this is usefull for creating data processing programs with many stored variables. See Info below
 
### More Info
 
2 variables

2 variables swapped

Eats 4 bytes of memory <-- no biggie


<span>             |<span>
---                |---
**Submitted On**   |
**By**             |[Bradley Liang](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByAuthor/bradley-liang.md)
**Level**          |Intermediate
**User Rating**    |4.8 (19 globes from 4 users)
**Compatibility**  |VB 4\.0 \(32\-bit\), VB 5\.0, VB 6\.0
**Category**       |[String Manipulation](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByCategory/string-manipulation__1-5.md)
**World**          |[Visual Basic](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByWorld/visual-basic.md)
**Archive File**   |[](https://github.com/Planet-Source-Code/bradley-liang-swap-variables-12-times-faster__1-8534/archive/master.zip)

### API Declarations

```
Private Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (Dest As Any, Source As Any, ByVal lNumBytes As Long)
```


### Source Code

```
Public Sub SwapStr(Var1 As String, Var2 As String)
' This is particularly useful in programs with lots of
' data analysis. Easily edited for any variant data
' manipulating. I'm currently using this coding and
' some vector codes to update my ThreeD Render Engine
' (http://www.planet-source-code.com/vb/scripts/ShowCode.asp?lngWId=1&txtCodeId=8426)
' a little advertising on my part =)...
' Using this routine is faster than
  ' sTmp = Var1
  ' Var1 = Var2
  ' Var2 = sTmp
' By a factor up 12 for really long values !!
Dim lSaveAddr As Long
' Save memory descriptor location for Var1
lSaveAddr = StrPtr(Var1)
' Copy memory descriptor of Var2 to Var1
CopyMemory ByVal VarPtr(Var1), ByVal VarPtr(Var2), 4
' Copy memory descriptor of saved Var1 to Var2
CopyMemory ByVal VarPtr(Var2), lSaveAddr, 4
'4 bytes is the size of one string. You may need to
'edit this coding a little in order to create memory
'efficient storage for different data types (i.e.
'user defined types).
End Sub
```


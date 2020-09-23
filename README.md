<div align="center">

## As Array got elements or not


</div>

### Description

Not many people know this one, an API that checks if an array as elements or not
 
### More Info
 


<span>             |<span>
---                |---
**Submitted On**   |
**By**             |[stephen whittle](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByAuthor/stephen-whittle.md)
**Level**          |Beginner
**User Rating**    |4.6 (32 globes from 7 users)
**Compatibility**  |VB 6\.0
**Category**       |[Windows API Call/ Explanation](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByCategory/windows-api-call-explanation__1-39.md)
**World**          |[Visual Basic](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByWorld/visual-basic.md)
**Archive File**   |[](https://github.com/Planet-Source-Code/stephen-whittle-as-array-got-elements-or-not__1-62272/archive/master.zip)





### Source Code

```
Private Declare Function SafeArrayGetDim Lib "oleaut32.dll" (ByRef saArray() As Any) As Long ----------------------------------
If SafeArrayGetDim(Your array) > 0 Then
 Debug.Print "Array as elements"
End If
```


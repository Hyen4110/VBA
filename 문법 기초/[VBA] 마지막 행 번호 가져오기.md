### VBA _ 마지막 행 번호 가져오기

```
Sub LastRow()
    r = Sheets(1).Range("A1048576").End(xlUp).Row
    MsgBox ("A열의 마지막 행은 " & r & "행입니다.")
End Sub
```

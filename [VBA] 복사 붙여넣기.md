
```
Sub OneCell()
    Range("A1:B10").Copy '복사명령 실행(Cut, Copy, Paste )
    Range("G1").Select '붙여넣을 범위 지정
    ActiveSheet.Paste '붙여넣기 지정
    Application.CutCopyMode = False

End Sub
```

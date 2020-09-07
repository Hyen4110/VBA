#### 시트목록 리스트 작성하기 

1. "시트목록" 이름으로 새로운 시트 생성
2. 아래 코드를 매크로에 추가하고 실행

```
  Dim Ws As Worksheet
  Dim ix As Integer, it As Integer
  Set Ws = Sheets("시트목록")

  With Ws
     it = Worksheets.Count
     .Columns("A:B").EntireColumn.ClearContents
     .Range("a1:b1") = Array("No", "시트명")

     For ix = 1 To it
        .Cells(ix + 1, 1) = ix
        .Cells(ix + 1, 2) = Sheets(ix).Name
    Next ix

    .Columns("A:B").EntireColumn.AutoFit

  End With

```

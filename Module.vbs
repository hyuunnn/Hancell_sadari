Sub moveLeg(ByRef name, ByRef i, ByRef legLength, ByRef CellIdx, ByRef startIdx)
  TimeNum = 0.1
  
  Do While 1
    ' 오른쪽에 있는 경우 +1 처리
    if Cells(i, CellIdx * 2 + 1) = "-" Then
      Cells(i, CellIdx * 2).Select
      Call WaitFor(TimeNum)
      CellIdx = Cellidx + 1
      Cells(i, CellIdx * 2).Select
      Call WaitFor(TimeNum)
      i = i + 1
          
    ' 왼쪽에 있는 경우 -1 처리 (B라인부터 시작하므로 A라인에 대한 예외처리 필요하지 않음)
    ElseIf Cells(i, CellIdx * 2 - 1) = "-" Then
      Cells(i, CellIdx * 2).Select
      Call WaitFor(TimeNum)
      CellIdx = CellIdx - 1
      Cells(i, CellIdx * 2).Select
      Call WaitFor(TimeNum)
      i = i + 1
              
    Else
      Cells(i, CellIdx * 2).Select
      Call WaitFor(TimeNum)
      i = i + 1
    End if
    
    ' 사다리 끝에 오면 종료
    if i = startIdx + legLength + 1 Then
      Cells(i, CellIdx * 2).Value = name
      Exit Do
    End if
  Loop
End Sub

' https://stackoverflow.com/questions/6960434/timing-delays-in-vba
Sub WaitFor(NumOfSeconds As Double)
  Dim SngSec as Double
  SngSec = Timer + NumOfSeconds
  Do While Timer < SngSec
    DoEvents
  Loop
End Sub

Sub checkWinner()
  ' 시작지점 (사다리타기가 시작되는 구간)
  startIdx = 10

  ' 전역변수 생성이 불가능하므로, 데이터를 엑셀에 저장한 후에 그 값을 가져와서 사용
  personLength = Range("D5:D5").Value
  legLength = Range("D6:D6").Value
  CellIdx = Range("J5:J5").Value

  If CellIdx > personLength Or CellIdx < 1 Then
    MsgBox("값이 잘못되었습니다.")
  Else
    ' 한셀에서는 ByRef만 사용 가능하여(ByRef는 참조타입) i를 사용할 때마다 초기화
    i = startidx
    name = Cells(i - 1, CellIdx * 2).Value
    Call moveLeg(name, i, legLength, CellIdx, startIdx)
  End If
End Sub

Sub checkWinnerAll()
  ' 시작지점 (사다리타기가 시작되는 구간)
  startIdx = 10

  ' 전역변수 생성이 불가능하므로, 데이터를 엑셀에 저장한 후에 그 값을 가져와서 사용
  personLength = Range("D5:D5").Value
  legLength = Range("D6:D6").Value

  For j = 1 To personLength
    ' 한셀에서는 ByRef만 사용 가능하여(ByRef는 참조타입) CellIdx, i를 사용할 때마다 초기화
    CellIdx = j
    i = startIdx
    name = Cells(i - 1, CellIdx * 2).Value
    Call moveLeg(name, i, legLength, CellIdx, startIdx)
  Next
End Sub

Sub initalize()
  ' 시작지점 (이름도 지정하기 위해서 시작 주소에 9를 사용)
  startIdx = 9
  personIdx = 1

  personLength_tmp = Range("D5:D5").Value
  legLength_tmp = Range("D6:D6").Value

  ' 새롭게 사다리를 생성하기 전에 이전에 저장된 사다리 삭제
  if personLength_tmp <> "" and legLength_tmp <> "" Then
    For i = 1 To personLength_tmp * 2 - 1
      ' 꽝, 당첨 부분과 이름이 입력되는 부분까지 삭제하기 위하여 +2 추가
      For j = 0 To legLength_tmp + 2
        With Cells(j + startIdx, i + 1)
          .Interior.Color = RGB(255,255,255)
          .Borders.LineStyle = xlContinuous
          .Borders.Color = RGB(190,190,190)
          .ClearContents
        End With
      Next
    Next
  End If

  personLength = InputBox("사다리 개수", "사다리타기", "10")
  legLength = InputBox("사다리 크기", "사다리타기", "20")

  ' 값이 존재하는 경우에만 동작
  If personLength <> "" and legLength <> "" Then
    Range("D5:D5").Value = personLength
    Range("D6:D6").Value = legLength

    ' 마지막 끝 열은 필요가 없기 때문에 -1 사용
    For i = 1 To personLength * 2 - 1
      If i mod 2 = 1 Then
        Cells(startIdx, i + 1).Value = personIdx
        personIdx = personIdx + 1
      End If
      
      For j = 1 To legLength
        ' 이쁘게 보이기 위하여 A열을 최대로 축소하였기 때문에 B부터 시작하여 1을 추가하였음
        If i mod 2 = 0 Then
          With Cells(startIdx + j, i + 1)
            .Interior.Color = RGB(255,204,204)
            .Value = "l"
          End With
        Else
          Cells(startIdx + j, i + 1).Value = "l"
        End If
      Next
    Next
  End If
End Sub

Sub legMakeAutomation()
  startIdx = 10

  personLength = Range("D5:D5").Value
  legLength = Range("D6:D6").Value

  ' C열에 해당하는 사다리부터 수정을 해야하기 때문에 2부터 시작
  For i = 2 To personLength
    ' startIdx 값인 10에 legLength만큼 for loop를 돌면 당첨, 꽝에 해당하는 부분까지 돌기 때문에 -1 사용
    For j = 0 To legLength - 1
      ' TODO: 랜덤으로 사다리 만드는 기능 개발
      Cells(startIdx + j, i * 2 - 1).Value = "l"
    Next
  Next
End Sub

Sub setRandomWinner()
  personLength = Range("D5:D5").Value
  legLength = Range("D6:D6").Value
  startIdx = 10

  Do While 1
    randomRnd = Int((personLength) * Rnd() + 1)
    if randomRnd <> 0 Then
      Exit Do
    End If
  Loop

  For i = 1 To personLength
    if i = randomRnd Then
      Cells(startIdx + legLength, i * 2).Value = "당점"
    Else
      Cells(startIdx + legLength, i * 2).Value = "꽝"
    End If
  Next
End Sub

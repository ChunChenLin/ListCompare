Sub filter()

    Dim range_A, range_B, count As Integer
    Dim s1, s2 As String
    Dim numOfSame As Integer
    Dim numOfNotSame As Integer
    
    'range_A = InputBox("請輸入老師網頁有幾筆資料")
    'range_B = InputBox("請輸入scopus有幾筆資料")
   
    range_A = 1
    range_B = 1
    
    '計算老師有幾筆資料
    While Worksheets("teacher").Cells(range_A + 1, 1) <> ""
        range_A = range_A + 1
    Wend
    
    '計算scopus有幾筆資料
    While Worksheets("scopus").Cells(range_B + 1, 1) <> ""
        range_B = range_B + 1
    Wend
   
    For i = 2 To range_A '判斷類型article, conference ...
        If Worksheets("teacher").Cells(i, 5) = "Journal" Then
            Worksheets("teacher").Cells(i, 12) = 1
        ElseIf Worksheets("teacher").Cells(i, 5) = "Conference" Then
            Worksheets("teacher").Cells(i, 12) = 2
        Else
            Worksheets("teacher").Cells(i, 12) = 3
        End If
    Next

    For i = 2 To range_A '一開始設為紫色
        'Set c = Worksheets("teacher").Cells(i, 2)
         Set c = Worksheets("teacher").Range("B" & i & ":G" & range_A & "")
            With c.Font
                .Color = -6279056
            End With
        Worksheets("teacher").Cells(i, 11).Value = 2
    Next
   
    
    For i = 2 To range_B '一開始設為綠色
        'Set c = Worksheets("scopus").Cells(i, 2)
         Set c = Worksheets("scopus").Range("A" & i & ":AP" & range_B & "")
            With c.Font
                .Color = -11489280
            End With
         Worksheets("scopus").Cells(i, 10).Value = 2
    Next
   
    numOfSame = 0
    For i = 2 To range_B '兩邊比對，若相同則改為紅色
        For j = 2 To range_A
            s1 = UCase(Worksheets("teacher").Cells(j, 3)) '解決大小寫問題
            s2 = UCase(Worksheets("scopus").Cells(i, 2))
            s1 = Replace(s1, " ", "") '解決多餘空格
            s2 = Replace(s2, " ", "")
            s1 = Replace(s1, ",", "")
            s2 = Replace(s2, ",", "")
            s1 = Replace(s1, ":", "")
            s2 = Replace(s2, ":", "")
            s1 = Replace(s1, "-", "")
            s2 = Replace(s2, "-", "")
            s1 = Replace(s1, ".", "")
            s2 = Replace(s2, ".", "")
            s1 = Replace(s1, "_", "")
            s2 = Replace(s2, "_", "")
            If s1 = s2 And _
            Worksheets("teacher").Cells(j, 4) = Worksheets("scopus").Cells(i, 3) Then
                numOfSame = numOfSame + 1
                
                Worksheets("teacher").Cells(j, 11).Value = 1
                Worksheets("scopus").Cells(i, 10).Value = 1
                'Set c = Worksheets("teacher").Cells(j, 2)
                Set c = Worksheets("teacher").Range("B" & j & ":G" & j & "")
                    With c.Font
                        .Color = -16776961
                    End With
                'Set c = Worksheets("scopus").Cells(i, 2)
                Set c = Worksheets("scopus").Range("A" & i & ":AP" & i & "")
                    With c.Font
                        .Color = -16776961
                    End With
            End If
        Next
    Next
    
    Columns("D:D").Select '排序年分
    ActiveWorkbook.Worksheets("teacher").Sort.SortFields.Clear
    ActiveWorkbook.Worksheets("teacher").Sort.SortFields.Add Key:=Range("D2"), _
        SortOn:=xlSortOnValues, order:=xlDescending, DataOption:=xlSortNormal
    With ActiveWorkbook.Worksheets("teacher").Sort
        .SetRange Range("B" & 2 & ":L" & range_A & "")
        .Header = xlNo
        .MatchCase = False
        .Orientation = xlTopToBottom
        .SortMethod = xlPinYin
        .Apply
    End With
    
    Columns("L:L").Select '排序類型
    ActiveWorkbook.Worksheets("teacher").Sort.SortFields.Clear
    ActiveWorkbook.Worksheets("teacher").Sort.SortFields.Add Key:=Range("L2"), _
        SortOn:=xlSortOnValues, order:=xlAscending, DataOption:=xlSortNormal
    With ActiveWorkbook.Worksheets("teacher").Sort
        .SetRange Range("B" & 2 & ":L" & range_A & "")
        .Header = xlNo
        .MatchCase = False
        .Orientation = xlTopToBottom
        .SortMethod = xlPinYin
        .Apply
    End With
    
    Columns("K:K").Select '排序老師狀況(顏色)
    ActiveWorkbook.Worksheets("teacher").Sort.SortFields.Clear
    ActiveWorkbook.Worksheets("teacher").Sort.SortFields.Add Key:=Range("K2"), _
        SortOn:=xlSortOnValues, order:=xlAscending, DataOption:=xlSortNormal
    With ActiveWorkbook.Worksheets("teacher").Sort
        .SetRange Range("B" & 2 & ":L" & range_A & "")
        .Header = xlNo
        .MatchCase = False
        .Orientation = xlTopToBottom
        .SortMethod = xlPinYin
        .Apply
    End With
    

    Columns("J:J").Select
    ActiveWorkbook.Worksheets("scopus").Sort.SortFields.Clear
    ActiveWorkbook.Worksheets("scopus").Sort.SortFields.Add Key:=Range("J1"), _
        SortOn:=xlSortOnValues, order:=xlDescending, DataOption:=xlSortNormal
    With ActiveWorkbook.Worksheets("scopus").Sort
        .SetRange Range("A2:AO75")
        .Header = xlNo
        .MatchCase = False
        .Orientation = xlTopToBottom
        .SortMethod = xlPinYin
        .Apply
    End With
  
    
    For i = 2 To range_A '清空類型紀錄(teacher)
        Worksheets("teacher").Cells(i, 12) = ""
        Worksheets("teacher").Cells(i, 11) = ""
    Next
    
 
    For i = 2 To range_B '清空類型紀錄(scopus)
        Worksheets("scopus").Cells(i, 10) = ""
    Next
 
    '新增工作表並更名
    Sheets.Add After:=Sheets(Sheets.Count)
    Sheets(Sheets.Count).Name = "Green"
    Sheets.Add After:=Sheets(Sheets.Count)
    Sheets(Sheets.Count).Name = "Red"
    Sheets.Add After:=Sheets(Sheets.Count)
    Sheets(Sheets.Count).Name = "Purple&Green"
    
    Worksheets("Green").Cells(1, 1).Value = "Authors"
    Worksheets("Green").Cells(1, 2).Value = "Title"
    Worksheets("Green").Cells(1, 3).Value = "Pub Year"
    Worksheets("Green").Cells(1, 4).Value = "Type"
    Worksheets("Red").Cells(1, 1).Value = "Authors"
    Worksheets("Red").Cells(1, 2).Value = "Title"
    Worksheets("Red").Cells(1, 3).Value = "Pub Year"
    Worksheets("Red").Cells(1, 4).Value = "Type"
    Worksheets("Purple&Green").Cells(1, 1).Value = "Authors"
    Worksheets("Purple&Green").Cells(1, 2).Value = "Title"
    Worksheets("Purple&Green").Cells(1, 3).Value = "Pub Year"
    Worksheets("Purple&Green").Cells(1, 4).Value = "Type"
    
    numOfNotSame = range_B - numOfSame
    For i=2 To numOfNotSame+2
        Worksheets("Green").Cells(i, 1).Value = Worksheets("scopus").Cells(i, 1).Value
        Worksheets("Green").Cells(i, 2).Value = Worksheets("scopus").Cells(i, 2).Value
        Worksheets("Green").Cells(i, 3).Value = Worksheets("scopus").Cells(i, 3).Value
        Worksheets("Green").Cells(i, 4).Value = Worksheets("scopus").Cells(i, 39).Value

        Worksheets("Purple&Green").Cells(i, 1).Value = Worksheets("scopus").Cells(i, 1).Value
        Worksheets("Purple&Green").Cells(i, 2).Value = Worksheets("scopus").Cells(i, 2).Value
        Worksheets("Purple&Green").Cells(i, 3).Value = Worksheets("scopus").Cells(i, 3).Value
        Worksheets("Purple&Green").Cells(i, 4).Value = Worksheets("scopus").Cells(i, 39).Value
    Next

    For i = 2 To numOfNotSame+2 'Green設為綠色
        'Set c = Worksheets("scopus").Cells(i, 2)
         Set c = Worksheets("Green").Range("A" & i & ":D" & numOfNotSame+2 & "")
            With c.Font
                .Color = -11489280
            End With
        Set d = Worksheets("Purple&Green").Range("A" & i & ":D" & numOfNotSame+2 & "")
            With d.Font
                .Color = -11489280
            End With
    Next
    
    j = 2
    For i=numOfNotSame+3 To range_B
        Worksheets("Red").Cells(j, 1).Value = Worksheets("scopus").Cells(i, 1).Value
        Worksheets("Red").Cells(j, 2).Value = Worksheets("scopus").Cells(i, 2).Value
        Worksheets("Red").Cells(j, 3).Value = Worksheets("scopus").Cells(i, 3).Value
        Worksheets("Red").Cells(j, 4).Value = Worksheets("scopus").Cells(i, 39).Value
        
        j = j + 1
    Next

    For i = 2 To numOfSame+2 'Red設為紅色
        'Set c = Worksheets("scopus").Cells(i, 2)
         Set c = Worksheets("Red").Range("A" & i & ":D" & numOfNotSame+2 & "")
            With c.Font
                .Color = -16776961
            End With
    Next

    j = numOfNotSame+3
    For i=numOfSame+1 To range_A
        Worksheets("Purple&Green").Cells(j, 1).Value = Worksheets("teacher").Cells(i, 2).Value
        Worksheets("Purple&Green").Cells(j, 2).Value = Worksheets("teacher").Cells(i, 3).Value
        Worksheets("Purple&Green").Cells(j, 3).Value = Worksheets("teacher").Cells(i, 4).Value
        Worksheets("Purple&Green").Cells(j, 4).Value = Worksheets("teacher").Cells(i, 5).Value
        
        j = j + 1
    Next

    For i = numOfNotSame+3 To numOfNotSame+3+range_A-numOfSame '紫色
        'Set c = Worksheets("scopus").Cells(i, 2)
        Set d = Worksheets("Purple&Green").Range("A" & i & ":D" & numOfNotSame+2 & "")
            With d.Font
                .Color = -6279056
            End With
    Next

    
    MsgBox ("紅色表示scopus有建檔" & vbCrLf & "綠色表示scopus有檔案 老師著作清單沒有列入" & vbCrLf & "紫色表示scopus沒有建檔")
    
End Sub





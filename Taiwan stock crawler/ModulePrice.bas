Attribute VB_Name = "Module1"
'main
Sub 按鈕1_Click()
    Dim now_sheet As String
    now_sheet = ActiveSheet.name
    Dim target_url As String
    target_url = getUrl()
On Error GoTo ErrorHandler
    getPrice target_url
    spanSetting
    Worksheets(now_sheet).Activate
    MsgBox ("更新完畢！")
Exit Sub
ErrorHandler:
    MsgBox ("資料不存在！")
    Worksheets(now_sheet).Activate
End Sub

Sub changeSheet(page_name As String)
    Dim sheet As Worksheet
    Dim sheet_exist As Boolean
    sheet_exist = False
    
    For Each sheet In Worksheets
        If sheet.name = page_name Then
            sheet_exist = True
            sheet.Select
            Exit For
        End If
    Next sheet

    If sheet_exist = False Then
        Worksheets.Add
        ActiveSheet.name = page_name
        'move sheet to last
        ActiveSheet.Move _
        After:=ActiveWorkbook.Sheets(ActiveWorkbook.Sheets.Count)
    End If
End Sub

Function getUrl() As String
    changeSheet "Date"
    Dim d As Date
    Dim string_y As String
    Dim string_m As String
    Dim string_d As String
    
    Cells(4, 1).Value = "今日日期"
    d = Date
    Cells(5, 1).Value = Year(d)
    Cells(5, 2).Value = Month(d)
    Cells(5, 3).Value = Day(d)
    
    Cells(1, 1).Value = "目標日期"
    On Error GoTo ErrorHandler
    d = DateSerial(CInt(Cells(2, 1).Value), CInt(Cells(2, 2).Value), CInt(Cells(2, 3).Value)) 'Type Error
    
    If d > Date Then 'after today
        d = Date
    ElseIf d < DateSerial(2004, 2, 11) Then 'no data
        d = Date
    End If
    'exclude weekend
    If Weekday(d) = 7 Then 'Satur
        d = DateAdd("d", -1, d)
    ElseIf Weekday(d) = 1 Then 'Sunday
        d = DateAdd("d", -2, d)
    End If
    Cells(2, 1).Value = Year(d)
    Cells(2, 2).Value = Month(d)
    Cells(2, 3).Value = Day(d)
    
    Cells(7, 1).Value = "股價來源"
    string_y = CStr(Year(d))
    If Month(d) < 10 Then
        string_m = "0" & CStr(Month(d))
    Else
        string_m = CStr(Month(d))
    End If
    If Day(d) < 10 Then
        string_d = "0" & CStr(Day(d))
    Else
        string_d = CStr(Day(d))
    End If
    getUrl = "https://www.twse.com.tw/exchangeReport/MI_INDEX?response=html&date=" & string_y & string_m & string_d & "&type=ALLBUT0999"
    Cells(8, 1).Value = getUrl
    
    Exit Function
    
ErrorHandler:
    d = Date
    MsgBox ("日期格式錯誤，跳回今日股價")
    Resume Next
End Function

Sub getPrice(url As String)
    changeSheet "Price"
    Dim priceXml As Object
    Dim priceHtml As Object
    
    Set priceXml = CreateObject("MSXML2.XMLHttp")
    Set priceHtml = CreateObject("HtmlFile")
    
    With priceXml
    
        .Open "GET", url, False
        .send
        priceHtml.body.innerhtml = .responsetext
        
        Set price_table = priceHtml.getelementsbytagname("table")(8) 'table
        Cells.Clear 'refresh
        i = 1
        For Each nRow In price_table.Rows
            j = 1
                For Each nCol In nRow.Cells
                    Cells(i, j) = nCol.innertext
                    j = j + 1
                Next nCol
            i = i + 1
        Next nRow
        
    End With
End Sub

Sub spanSetting()
    Worksheets("Date").Activate
    Cells(10, 1).Value = "輸入代碼"
    Cells(10, 2).Value = "公司名稱"
    Cells(10, 3).Value = "收盤價"
    Cells(10, 4).Value = "row"
    Cells(10, 5).Value = "column"
    
    If Cells(11, 1).Value < 50 Then '預設台灣50
        Cells(11, 1).Value = 50
    End If
    
    '以下 code 自錄製巨集
    Range("D11").Select
    ActiveCell.FormulaR1C1 = "=MATCH(RC[-3],Price!C[-3],0)"
    Range("E11").Select
    ActiveCell.FormulaR1C1 = "=MATCH(R10C3,Price!R3C1:R3C20,0)"
    Range("B11").Select
    ActiveCell.FormulaR1C1 = "=INDEX(Price!C[-1]:C[18],RC[2],2)"
    Range("C11").Select
    ActiveCell.FormulaR1C1 = "=INDEX(Price!C[-2]:C[17],RC[1],RC[2])"
    
End Sub

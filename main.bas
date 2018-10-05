Sub Main()

    List_head
    
    Dim strText As String
    Dim arrRow, arrCell
    Dim i As Long, j As Long, n As Long
    Dim arrColumn
    Dim arrData(1 To 1000, 1 To 10)
    On Error Resume Next
    
With CreateObject("WinHttp.WinHttpRequest.5.1")
        .Open "GET", "https://movie.douban.com/top250"
 '       .setRequestHeader "Cookie", ""
        .Send
        strText = .responsetext
        Debug.Print strText
    End With


      arrRow = Split(strText, "<img width=")
'    Range("a255").Resize(25, 1) = Application.Transpose(arrRow)
       For n = 1 To 25
       Cells(n + 1, 1).Value = Split(arrRow(n), """")(3)
       Cells(n + 1, 2).Value = Split(Split(Split(Split(arrRow(n), "导演: ")(1), ">")(1), " ")(28), "&")(0)
       Cells(n + 1, 3).Value = Split(Split(Split(Split(arrRow(n), "导演: ")(1), ">")(1), ";")(2), "&")(0)
       Cells(n + 1, 4).Value = Left(Split(Split(arrRow(n), "导演: ")(1), "&")(0), 50)
       Cells(n + 1, 5).Value = Split(Split(Split(arrRow(n), "导演: ")(1), "主演: ")(1), "<")(0)
       Cells(n + 1, 6).Value = Split(Split(Split(arrRow(n), "v:average")(1), ">")(1), "<")(0)
       Cells(n + 1, 7).Value = Split(Split(Split(Split(arrRow(n), "v:average")(1), ">")(5), "<")(0), "人")(0)
       Cells(n + 1, 8).Value = Split(arrRow(n), """")(13)
       Next
      
  For nm = 25 To 225 Step 25

      
    With CreateObject("WinHttp.WinHttpRequest.5.1")
        .Open "GET", "https://movie.douban.com/top250?start=" & nm & "&filter="
 '       .setRequestHeader "Cookie", ""
        .Send
        strText = .responsetext
        Debug.Print strText
    End With
    

      arrRow = Split(strText, "<img width=")
       For n = 1 To 25
       Cells(n + nm + 1, 1).Value = Split(arrRow(n), """")(3)
       Cells(n + nm + 1, 2).Value = Split(Split(Split(Split(arrRow(n), "导演: ")(1), ">")(1), " ")(28), "&")(0)
       Cells(n + nm + 1, 3).Value = Split(Split(Split(Split(arrRow(n), "导演: ")(1), ">")(1), ";")(2), "&")(0)
       Cells(n + nm + 1, 4).Value = Left(Split(Split(arrRow(n), "导演: ")(1), "&")(0), 50)
       Cells(n + nm + 1, 5).Value = Split(Split(Split(arrRow(n), "导演: ")(1), "主演: ")(1), "<")(0)
       Cells(n + nm + 1, 6).Value = Split(Split(Split(arrRow(n), "v:average")(1), ">")(1), "<")(0)
       Cells(n + nm + 1, 7).Value = Split(Split(Split(Split(arrRow(n), "v:average")(1), ">")(5), "<")(0), "人")(0)
       Cells(n + nm + 1, 8).Value = Split(arrRow(n), """")(13)
       Next
        
    Next
    
      
End Sub
Sub List_head()

'
    Range("A1").Select
    ActiveCell.FormulaR1C1 = "电影"
    Range("B1").Select
    ActiveCell.FormulaR1C1 = "年份"
    Range("C1").Select
    ActiveCell.FormulaR1C1 = "国家"
    Range("D1").Select
    ActiveCell.FormulaR1C1 = "导演"
    Range("E1").Select
    ActiveCell.FormulaR1C1 = "主演"
    Range("F1").Select
    ActiveCell.FormulaR1C1 = "评分"
    Range("G1").Select
    ActiveCell.FormulaR1C1 = "评分人数"
    Range("H1").Select
    ActiveCell.FormulaR1C1 = "豆瓣地址"
    Range("A1:H1").Select
    Selection.Font.Bold = True
    Columns("G:G").ColumnWidth = 11
    Columns("H:H").ColumnWidth = 47
    Columns("F:F").ColumnWidth = 7
    Columns("E:E").ColumnWidth = 49
    Columns("D:D").ColumnWidth = 66
    Columns("C:C").ColumnWidth = 29
    Columns("B:B").ColumnWidth = 15
    Columns("A:A").ColumnWidth = 25
    Range("A1:H251").Select
    With Selection.Font
        .Name = "微软雅黑"
        .Strikethrough = False
        .Superscript = False
        .Subscript = False
        .OutlineFont = False
        .Shadow = False
        .Underline = xlUnderlineStyleNone
        .ThemeColor = xlThemeColorLight1
        .TintAndShade = 0
        .ThemeFont = xlThemeFontNone
    End With
    With Selection.Font
        .Name = "微软雅黑"
        .Size = 11
        .Strikethrough = False
        .Superscript = False
        .Subscript = False
        .OutlineFont = False
        .Shadow = False
        .Underline = xlUnderlineStyleNone
        .ThemeColor = xlThemeColorLight1
        .TintAndShade = 0
        .ThemeFont = xlThemeFontNone
    End With
    With Selection
        .HorizontalAlignment = xlGeneral
        .VerticalAlignment = xlCenter
        .Orientation = 0
        .AddIndent = False
        .IndentLevel = 0
        .ShrinkToFit = False
        .ReadingOrder = xlContext
        .MergeCells = False
    End With
    With Selection
        .HorizontalAlignment = xlCenter
        .VerticalAlignment = xlCenter
        .Orientation = 0
        .AddIndent = False
        .IndentLevel = 0
        .ShrinkToFit = False
        .ReadingOrder = xlContext
        .MergeCells = False
    End With
    Range("A1:H1").Select
    With Selection.Font
        .Name = "微软雅黑"
        .Size = 14
        .Strikethrough = False
        .Superscript = False
        .Subscript = False
        .OutlineFont = False
        .Shadow = False
        .Underline = xlUnderlineStyleNone
        .ThemeColor = xlThemeColorLight1
        .TintAndShade = 0
        .ThemeFont = xlThemeFontNone
    End With
    Range("A1").Select
        Range("A1:H251").Select
    Selection.Borders(xlDiagonalDown).LineStyle = xlNone
    Selection.Borders(xlDiagonalUp).LineStyle = xlNone
    With Selection.Borders(xlEdgeLeft)
        .LineStyle = xlContinuous
        .ColorIndex = 0
        .TintAndShade = 0
        .Weight = xlThin
    End With
    With Selection.Borders(xlEdgeTop)
        .LineStyle = xlContinuous
        .ColorIndex = 0
        .TintAndShade = 0
        .Weight = xlThin
    End With
    With Selection.Borders(xlEdgeBottom)
        .LineStyle = xlContinuous
        .ColorIndex = 0
        .TintAndShade = 0
        .Weight = xlThin
    End With
    With Selection.Borders(xlEdgeRight)
        .LineStyle = xlContinuous
        .ColorIndex = 0
        .TintAndShade = 0
        .Weight = xlThin
    End With
    With Selection.Borders(xlInsideVertical)
        .LineStyle = xlContinuous
        .ColorIndex = 0
        .TintAndShade = 0
        .Weight = xlThin
    End With
    With Selection.Borders(xlInsideHorizontal)
        .LineStyle = xlContinuous
        .ColorIndex = 0
        .TintAndShade = 0
        .Weight = xlThin
    End With
    Range("A1").Select
End Sub

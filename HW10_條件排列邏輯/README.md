# 這是個依照題目條件排列的VBA實作
請打開excel >> 點選開發人員 >> 點選Visual Basic >> 觀看程式碼與工作表上結果呈現

---

> 本題中共出現了這些選項(未經排列，直接分類隨機列出而已)
\   |  房客1  |   房客2  |   房客3  |   房客4  |   房客5
------------ | ------------- | ------------ | ------------- | ------------- | -------------
顏色： | 黃色  | 藍色  | 白色  | 綠色  | 紅色
國籍： | 挪威  | 泰國  | 美國  | 中國  | 英國 
飲料： | 橘子  | 茶    |  牛奶  | 咖啡  | 水 
球類： | 籃球  | 保齡  | 羽球  | 手球  | 桌球 
寵物： | 狐狸  | 斑馬  | 狗     |  羊    |   蝸牛 

> 建立五個陣列分別儲存五種屬性(cc(1) = “紅色”、cc(2) = “綠色”、cc(3) = “白色“……各陣列代表的性質固定)。建立五個for迴圈(各120種「位置」來篩選)相包，一層一層篩選出符合各個屬性中條件的結果(陣列性質固定，所以篩選出符合所有條件的「位置」)，最後把陣列一個一個輸出到符合所有條件的各個位置。
我的陣列性質固定，120種組合代表的是各性質填入的「位置」，故我的5個for迴圈裏面都會先包「將120種組合的五個數字分別填入五個位置的程式碼」
例如：
For c = 1 To 120
c1 = Mid(a(c), 1, 1)
c2 = Mid(a(c), 2, 1)
c3 = Mid(a(c), 3, 1)
c4 = Mid(a(c), 4, 1)
c5 = Mid(a(c), 5, 1)
……(接著依照題幹寫判斷條件)


> 編碼表： 
顏色 | 國籍  | 飲料  | 運動  | 寵物 
------------ | ------------- | ------------ | ------------- | -------------
cc(1) = "紅色“ | 	nn(1) = "英國“ | 	bb(1) = "水“	 | ss(1) = "羽球“	 | pp(1) = "狐狸"
cc(2) = "綠色“ | 	nn(2) = "挪威“ | 	bb(2) = "咖啡“	 | ss(2) = "手球“	 | pp(2) = "羊"
cc(3) = "白色“ | 	nn(3) = "中國“ | 	bb(3) = "橘子“	 | ss(3) = "籃球“	 | pp(3) = "斑馬"
cc(4) = "藍色“ | 	nn(4) = "泰國“ | 	bb(4) = "茶“	 | ss(4) = "保齡“	 | pp(4) = "狗"
cc(5) = "黃色“ | 	nn(5) = "美國“ | 	bb(5) = "牛奶“	 | ss(5) = "桌球“	 | pp(5) = "蝸牛“

---

以下為程式碼：

```VBA
Option Explicit
Sub 實習十嘗試()

Dim a(1 To 120) As Single
Dim i!, j!, stTimer As Single
Dim c%, n%, b%, s%, p As Integer
Dim c1%, c2%, c3%, c4%, c5 As Integer
Dim n1%, n2%, n3%, n4%, n5 As Integer
Dim b1%, b2%, b3%, b4%, b5 As Integer
Dim s1%, s2%, s3%, s4%, s5 As Integer
Dim p1%, p2%, p3%, p4%, p5 As Integer
Dim k%, m As Integer
Dim cc$(1 To 5), nn$(1 To 5), bb$(1 To 5), ss$(1 To 5), pp(1 To 5) As String


ActiveSheet.Cells.Clear
stTimer = Timer

'寫出12345不重複排列的120種結果，並存到陣列a()中

j = 1
For i = 12345 To 54321
    If InStr(i, "0") = 0 And InStr(i, "9") = 0 And InStr(i, "8") = 0 And InStr(i, "7") = 0 And InStr(i, "6") = 0 Then
        If Mid(i, 1, 1) <> Mid(i, 2, 1) And Mid(i, 1, 1) <> Mid(i, 3, 1) And Mid(i, 1, 1) <> Mid(i, 4, 1) And Mid(i, 1, 1) <> Mid(i, 5, 1) _
        And Mid(i, 2, 1) <> Mid(i, 3, 1) And Mid(i, 2, 1) <> Mid(i, 4, 1) And Mid(i, 2, 1) <> Mid(i, 5, 1) _
        And Mid(i, 3, 1) <> Mid(i, 4, 1) And Mid(i, 3, 1) <> Mid(i, 5, 1) _
        And Mid(i, 4, 1) <> Mid(i, 5, 1) Then
            a(j) = i
            j = j + 1
        End If
    End If
Next

cc(1) = "紅色"
cc(2) = "綠色"
cc(3) = "白色"
cc(4) = "藍色"
cc(5) = "黃色"
'房子顏色   c1=紅   c2=綠   c3=白   c4=藍   c5=黃
nn(1) = "英國"
nn(2) = "挪威"
nn(3) = "中國"
nn(4) = "泰國"
nn(5) = "美國"
'國籍       n1=英國 n2=挪威 n3=中國 n4=泰國 n5=美國
bb(1) = "水"
bb(2) = "咖啡"
bb(3) = "橘子"
bb(4) = "茶"
bb(5) = "牛奶"
'飲料       b1=水   b2=咖啡 b3=橘子 b4=茶   b5=牛奶
ss(1) = "羽球"
ss(2) = "手球"
ss(3) = "籃球"
ss(4) = "保齡"
ss(5) = "桌球"
'球類運動   s1=羽球 s2=手球 s3=籃球 s4=保齡 s5=桌球
pp(1) = "狐狸"
pp(2) = "羊"
pp(3) = "斑馬"
pp(4) = "狗"
pp(5) = "蝸牛"
'寵物       p1=狐狸 p2=羊   p3=斑馬 p4=狗   p5=蝸牛

For n = 1 To 120
    
    n1 = Mid(a(n), 1, 1)
    n2 = Mid(a(n), 2, 1)
    n3 = Mid(a(n), 3, 1)
    n4 = Mid(a(n), 4, 1)
    n5 = Mid(a(n), 5, 1)
    
    If n2 = 1 Then '(11)挪威人第1棟
    
    
        For b = 1 To 120
        
        b1 = Mid(a(b), 1, 1)
        b2 = Mid(a(b), 2, 1)
        b3 = Mid(a(b), 3, 1)
        b4 = Mid(a(b), 4, 1)
        b5 = Mid(a(b), 5, 1)
        
        If b5 = 3 Then '(10)喝牛奶的住第3棟
        If b4 = n4 Then '(6)泰國人喝茶

            For c = 1 To 120
            
            c1 = Mid(a(c), 1, 1)
            c2 = Mid(a(c), 2, 1)
            c3 = Mid(a(c), 3, 1)
            c4 = Mid(a(c), 4, 1)
            c5 = Mid(a(c), 5, 1)
            
            If n2 = c4 + 1 Or n2 = c4 - 1 Then '(14)挪威人住藍色房子的旁邊    (第2棟
            If n1 = c1 Then '(3)英國人住紅色
            If b2 = c2 Then '(5)住綠色房子的人喝咖啡
            If c2 - 1 = c3 Then '(7)綠色房子左邊是白色房子
        
                For s = 1 To 120
                
                s1 = Mid(a(s), 1, 1)
                s2 = Mid(a(s), 2, 1)
                s3 = Mid(a(s), 3, 1)
                s4 = Mid(a(s), 4, 1)
                s5 = Mid(a(s), 5, 1)
                
                If s2 = n3 Then '(12)中國人打手球
                If s3 = b3 Then '(13)打籃球的人喝橘子汁
                If s1 = c5 Then '(9)住黃色房子的人打羽球
                 
                    For p = 1 To 120
                    
                    p1 = Mid(a(p), 1, 1)
                    p2 = Mid(a(p), 2, 1)
                    p3 = Mid(a(p), 3, 1)
                    p4 = Mid(a(p), 4, 1)
                    p5 = Mid(a(p), 5, 1)
                    
                    If p4 = n5 Then '(4)養狗的是美國人
                    If p5 = s5 Then '(8)打桌球的人養蝸牛
                    If p1 + 1 = s4 Or p1 - 1 = s4 Then '(15)打保齡球的住在養狐狸的隔壁
                    If p2 + 1 = s1 Or p2 - 1 = s1 Then '(16)打羽球的住在養羊的隔壁
                    
                        'If a(c) = a(n) And a(n) = a(b) And a(b) = a(s) And a(s) = a(p) Then
                            
                            
                                ActiveSheet.Cells(3, 2 + c1) = cc(1)
                                ActiveSheet.Cells(3, 2 + c2) = cc(2)
                                ActiveSheet.Cells(3, 2 + c3) = cc(3)
                                ActiveSheet.Cells(3, 2 + c4) = cc(4)
                                ActiveSheet.Cells(3, 2 + c5) = cc(5)
                                
                                ActiveSheet.Cells(4, 2 + n1) = nn(1)
                                ActiveSheet.Cells(4, 2 + n2) = nn(2)
                                ActiveSheet.Cells(4, 2 + n3) = nn(3)
                                ActiveSheet.Cells(4, 2 + n4) = nn(4)
                                ActiveSheet.Cells(4, 2 + n5) = nn(5)
                                
                                ActiveSheet.Cells(5, 2 + b1) = bb(1)
                                ActiveSheet.Cells(5, 2 + b2) = bb(2)
                                ActiveSheet.Cells(5, 2 + b3) = bb(3)
                                ActiveSheet.Cells(5, 2 + b4) = bb(4)
                                ActiveSheet.Cells(5, 2 + b5) = bb(5)
                            
                                ActiveSheet.Cells(6, 2 + s1) = ss(1)
                                ActiveSheet.Cells(6, 2 + s2) = ss(2)
                                ActiveSheet.Cells(6, 2 + s3) = ss(3)
                                ActiveSheet.Cells(6, 2 + s4) = ss(4)
                                ActiveSheet.Cells(6, 2 + s5) = ss(5)
                                
                                ActiveSheet.Cells(7, 2 + p1) = pp(1)
                                ActiveSheet.Cells(7, 2 + p2) = pp(2)
                                ActiveSheet.Cells(7, 2 + p3) = pp(3)
                                ActiveSheet.Cells(7, 2 + p4) = pp(4)
                                ActiveSheet.Cells(7, 2 + p5) = pp(5)
                            
                            
                        'End If
                
                    End If
                    End If
                    End If
                    End If
                    Next
                    
                End If
                End If
                End If
                Next
                
            End If
            End If
            End If
            End If
            Next
            
        End If
        End If
        Next
        
    End If
    Next


ActiveSheet.Cells(2, 2) = Timer - stTimer

' 框線 巨集
    Range("C3:G7").Select
    Selection.Borders(xlDiagonalDown).LineStyle = xlNone
    Selection.Borders(xlDiagonalUp).LineStyle = xlNone
    With Selection.Borders(xlEdgeLeft)
        .LineStyle = xlContinuous
        .ColorIndex = 0
        .TintAndShade = 0
        .Weight = xlMedium
    End With
    With Selection.Borders(xlEdgeTop)
        .LineStyle = xlContinuous
        .ColorIndex = 0
        .TintAndShade = 0
        .Weight = xlMedium
    End With
    With Selection.Borders(xlEdgeBottom)
        .LineStyle = xlContinuous
        .ColorIndex = 0
        .TintAndShade = 0
        .Weight = xlMedium
    End With
    With Selection.Borders(xlEdgeRight)
        .LineStyle = xlContinuous
        .ColorIndex = 0
        .TintAndShade = 0
        .Weight = xlMedium
    End With
    Selection.Borders(xlInsideVertical).LineStyle = xlNone
    Selection.Borders(xlInsideHorizontal).LineStyle = xlNone
    Range("D3:D7").Select
    Selection.Borders(xlDiagonalDown).LineStyle = xlNone
    Selection.Borders(xlDiagonalUp).LineStyle = xlNone
    With Selection.Borders(xlEdgeLeft)
        .LineStyle = xlContinuous
        .ColorIndex = 0
        .TintAndShade = 0
        .Weight = xlMedium
    End With
    With Selection.Borders(xlEdgeTop)
        .LineStyle = xlContinuous
        .ColorIndex = 0
        .TintAndShade = 0
        .Weight = xlMedium
    End With
    With Selection.Borders(xlEdgeBottom)
        .LineStyle = xlContinuous
        .ColorIndex = 0
        .TintAndShade = 0
        .Weight = xlMedium
    End With
    With Selection.Borders(xlEdgeRight)
        .LineStyle = xlContinuous
        .ColorIndex = 0
        .TintAndShade = 0
        .Weight = xlMedium
    End With
    Selection.Borders(xlInsideVertical).LineStyle = xlNone
    Selection.Borders(xlInsideHorizontal).LineStyle = xlNone
    Range("F3:F7").Select
    Selection.Borders(xlDiagonalDown).LineStyle = xlNone
    Selection.Borders(xlDiagonalUp).LineStyle = xlNone
    With Selection.Borders(xlEdgeLeft)
        .LineStyle = xlContinuous
        .ColorIndex = 0
        .TintAndShade = 0
        .Weight = xlMedium
    End With
    With Selection.Borders(xlEdgeTop)
        .LineStyle = xlContinuous
        .ColorIndex = 0
        .TintAndShade = 0
        .Weight = xlMedium
    End With
    With Selection.Borders(xlEdgeBottom)
        .LineStyle = xlContinuous
        .ColorIndex = 0
        .TintAndShade = 0
        .Weight = xlMedium
    End With
    With Selection.Borders(xlEdgeRight)
        .LineStyle = xlContinuous
        .ColorIndex = 0
        .TintAndShade = 0
        .Weight = xlMedium
    End With
    Selection.Borders(xlInsideVertical).LineStyle = xlNone
    Selection.Borders(xlInsideHorizontal).LineStyle = xlNone

End Sub
```

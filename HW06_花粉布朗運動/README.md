# 這是個模擬 1D / 2D 布朗運動的VBA實作
請打開excel >> 點選開發人員 >> 點選Visual Basic >> 觀看程式碼與工作表上結果呈現

---
以下為程式碼：

# 1. 一維布朗運動
> 列出此模擬(即一千個顆粒，各跳一萬次)的結果，按模擬次數序號、跳一萬次後之位置、  
>     跳一萬次後距原點之距離(絕對值)、跑程式時間等四項列表。
```VBA
Option Explicit
Sub random_walk_1D()
Dim i&, j&, sum_position&, abs_sum_position&, sum_RSquare&, RBar&, d As Long
Dim position(1 To 1000) As Long
Dim stTimer!, avg_abs_position!, avg_position!, k As Single

ActiveSheet.Cells.Clear

RandomizeX

d = 0
Do
stTimer = Timer

    For i = 1 To 1000
        position(i) = 0
        For j = 1 To 10000
            If RndX > 0.5 Then
                position(i) = position(i) + 1
                Else
                position(i) = position(i) - 1
            End If
        Next
    Next
    
    sum_position = 0
    abs_sum_position = 0
    avg_position = 0
    avg_abs_position = 0
    'sum_RSquare = 0
    
    For i = 1 To 1000
        sum_position = sum_position + position(i)
        abs_sum_position = abs_sum_position + Abs(position(i))
        'sum_RSquare = sum_RSquare + (Abs(position(i))) ^ 2
    Next
    avg_position = sum_position / 1000
    avg_abs_position = abs_sum_position / 1000
    'RBar = sum_RSquare / 1000
    ActiveWorkbook.Worksheets("第三頁").Select
    Cells(8, 2 + d * 4) = "No."
    Cells(8, 3 + d * 4) = "Position"
    Cells(8, 4 + d * 4) = "Abs_Position"
    For i = 9 To 1008
        Cells(i, 2 + d * 4) = i - 8
        Cells(i, 3 + d * 4) = position(i - 8)
        Cells(i, 4 + d * 4) = Abs(position(i - 8))
    Next
    Cells(6, 3 + d * 4) = "Avg:"
    Cells(7, 3 + d * 4) = avg_position
    Cells(6, 4 + d * 4) = "Abs.Avg:"
    Cells(7, 4 + d * 4) = avg_abs_position
    'ActiveSheet.Cells(1006, 2).Value = "方均根："
    'ActiveSheet.Cells(1006, 3).Value = Sqr(RBar)
    Cells(5, 2 + d * 4) = "跑程式時間：" & Timer - stTimer & "秒"
    Cells(4, 2 + d * 4) = "第" & Str(d + 1) & "次執行"

    d = d + 1
    
Loop Until d = 10

End Sub
```

---

> 算出這模擬共1000個顆粒跳動後的總平均位置、總平均距離、距離的方均根。

```VBA
Option Explicit
Sub radom_walk_relation()

Dim i&, j&, x%, sum_position&, abs_sum_position&
Dim avg_position&, avg_abs_position&
Dim position(1 To 1000) As Long
Dim k As Long, sumsqr As Long
Dim time!

x = 0
Do

time = Timer

RandomizeX
For i = 1 To 1000
    position(i) = 0
    For j = 1 To 1000 + x * 100
        If RndX > 0.5 Then
        position(i) = position(i) + 1
        Else
        position(i) = position(i) - 1
        End If
    Next j
Next i
    
    sum_position = 0
    abs_sum_position = 0
    sumsqr = 0
        
For i = 1 To 1000
    sum_position = sum_position + position(i)
    abs_sum_position = abs_sum_position + Abs(position(i))
    sumsqr = sumsqr + position(i) ^ 2
Next i
    
    avg_position = sum_position / 1000
    avg_abs_position = abs_sum_position / 1000
    k = (sumsqr / 1000) ^ 0.5
    ActiveWorkbook.Worksheets("工作表2").Select

   
    ActiveSheet.Cells(3, 2) = "跳動次數"
    ActiveSheet.Cells(3, 3) = "總平均位置"
    ActiveSheet.Cells(3, 4) = "總平均距離"
    ActiveSheet.Cells(3, 5) = "距離方均根"
    ActiveSheet.Cells(3, 6).Value = "執行時間"
    
    ActiveSheet.Cells(4 + x, 2) = x * 100 + 1000
    ActiveSheet.Cells(4 + x, 3) = avg_position
    ActiveSheet.Cells(4 + x, 4) = avg_abs_position
    ActiveSheet.Cells(4 + x, 5) = k
    ActiveSheet.Cells(4 + x, 6) = Timer - time

x = x + 1
Loop Until x > 90

End Sub
```

---

> 每次跳動有三種可能(向左、不動、向右)且機率軍相等，position平均一樣在0附近，  
> 但平均Abs_position從80左右下降到65左右。

```VBA
Option Explicit
Sub random_walk_1D2()
Dim i&, j&, sum_position&, abs_sum_position&, sum_RSquare&, RBar&, d As Long
Dim position(1 To 1000) As Long
Dim stTimer!, k!, avg_position!, avg_abs_position!, position_2(1 To 10000) As Single

ActiveSheet.Cells.Clear

RandomizeX

d = 0
Do
stTimer = Timer

sum_position = 0
abs_sum_position = 0
avg_position = 0
avg_abs_position = 0
sum_RSquare = 0

    For i = 1 To 1000
        position(i) = 0
        For j = 1 To 10000
            position_2(j) = RndX
            If position_2(j) > (2 / 3) Then
                position(i) = position(i) + 1
            ElseIf position_2(j) < (1 / 3) Then
                position(i) = position(i) - 1
            End If
        Next
    Next
    
    For i = 1 To 1000
        sum_position = sum_position + position(i)
        abs_sum_position = abs_sum_position + Abs(position(i))
        sum_RSquare = sum_RSquare + (Abs(position(i))) ^ 2
    Next
    avg_position = sum_position / 1000
    avg_abs_position = abs_sum_position / 1000
    RBar = sum_RSquare / 1000
    ActiveWorkbook.Worksheets("工作表3").Select
    ActiveSheet.Cells(8, 2 + d * 4).Value = "No."
    ActiveSheet.Cells(8, 3 + d * 4).Value = "Position"
    ActiveSheet.Cells(8, 4 + d * 4).Value = "Abs_Position"
    For i = 9 To 1008
        ActiveSheet.Cells(i, 2 + d * 4).Value = i - 8
        ActiveSheet.Cells(i, 3 + d * 4).Value = position(i - 8)
        ActiveSheet.Cells(i, 4 + d * 4).Value = Abs(position(i - 8))
    Next
    ActiveSheet.Cells(6, 3 + d * 4).Value = "Avg:"
    ActiveSheet.Cells(7, 3 + d * 4).Value = avg_position
    ActiveSheet.Cells(6, 4 + d * 4).Value = "Abs.Avg:"
    ActiveSheet.Cells(7, 4 + d * 4).Value = avg_abs_position
    'ActiveSheet.Cells(1011, 2).Value = "方均根："
    'ActiveSheet.Cells(1011, 3).Value = Sqr(RBar)
    ActiveSheet.Cells(5, 2 + d * 4).Value = "跑程式時間："
    ActiveSheet.Cells(5, 4 + d * 4).Value = Timer - stTimer
    ActiveSheet.Cells(4, 2 + d * 4).Value = "第" & Str(d + 1) & "次執行"
    
    d = d + 1
    
 Loop Until d = 10
    
End Sub
```

---

# 2. 二維布朗運動
> 將這十次模擬(每次有1000顆)的平均結果列表，包括
> 1. 模擬序數、
> 2. 跳一萬次後與原點之平均距離(絕對值)、
> 3. 與原點之平均方位(角度)、
> 4. x座標、
> 5. y座標、 
> 6. 和跑程式時間
> 等六項。

```VBA
Option Explicit
Sub randomwalk_2D()
Dim i&, j&, x%
Dim sum_posx&, sum_posy&, sum_angle!
Dim avg_posx!, avg_posy!, avg_angle!, avg_distance!
Dim posx&(1 To 1000), posy&(1 To 1000), angle!(1 To 1000)
Dim pi!, time!

x = 0
Do
time = Timer

pi = 3.1415926

ReDim pos(10000)
RandomizeX
For i = 1 To 1000
    posx(i) = 0
    posy(i) = 0
        For j = 1 To 10000
            pos(j) = 2 * pi * RndX
            posx(i) = posx(i) + Cos(pos(j))
            posy(i) = posy(i) + Sin(pos(j))
        Next j
        
        If posx(i) > 0 And posy(i) >= 0 Then
        angle(i) = Atn(posy(i) / posx(i)) * 180 / pi
        
        ElseIf posx(i) < 0 And posy(i) >= 0 Then
        angle(i) = Atn(posy(i) / posx(i)) * 180 / pi + 180
        
        ElseIf posx(i) < 0 And posy(i) < 0 Then
        angle(i) = Atn(posy(i) / posx(i)) * 180 / pi + 180
        
        ElseIf posx(i) > 0 And posy(i) < 0 Then
        angle(i) = Atn(posy(i) / posx(i)) * 180 / pi + 360
        
        ElseIf posx(i) = 0 And posy(i) > 0 Then
        angle(i) = 90
        
        ElseIf posx(i) = 0 And posy(i) < 0 Then
        angle(i) = 270
        
       End If
Next i
    
    
    sum_posx = 0
    sum_posy = 0
    sum_angle = 0
    
For i = 1 To 1000
        
sum_posx = sum_posx + posx(i)
sum_posy = sum_posy + posy(i)
sum_angle = sum_angle + angle(i)

Next i
    
    avg_posx = sum_posx / 1000
    avg_posy = sum_posy / 1000
    avg_angle = sum_angle / 1000
    avg_distance = Sqr(avg_posx ^ 2 + avg_posy ^ 2)
    
    ActiveWorkbook.Worksheets("第三頁").Select
    ActiveSheet.Cells(3, 2 + x * 7).Value = "序號"
    ActiveSheet.Cells(3, 3 + x * 7).Value = "x座標"
    ActiveSheet.Cells(3, 4 + x * 7).Value = "y座標"
    ActiveSheet.Cells(3, 5 + x * 7).Value = "方位"
    ActiveSheet.Cells(3, 6 + x * 7).Value = "平均距離"
    ActiveSheet.Cells(3, 7 + x * 7).Value = "平均方位"

For i = 4 To 1003
        ActiveSheet.Cells(i, 2 + x * 7) = i - 3
        ActiveSheet.Cells(i, 3 + x * 7) = posx(i - 3)
        ActiveSheet.Cells(i, 4 + x * 7) = posy(i - 3)
        ActiveSheet.Cells(i, 5 + x * 7) = angle(i - 3)
        
Next i
    
    ActiveSheet.Cells(4, 6 + x * 7) = avg_distance
    ActiveSheet.Cells(4, 7 + x * 7) = avg_angle
    ActiveSheet.Cells(2, 2 + x * 7).Value = "執行第" & Str(x + 1) & "次，執行時間" & Timer - time & "秒"

x = x + 1
Loop Until x >= 10
End Sub
```

---

> 分析平均距離、平均方位角這兩個值，與跳動次數的關係，
> 發現平均距離皆位於0左右，平均方位角則在180左右。

```VBA
Option Explicit
Sub random_walk_relation()
Dim i&, j&, x%
Dim sum_posx&, sum_posy&, sum_angle!, sum_distance!
Dim avg_posx!, avg_posy!, avg_angle!, avg_distance!
Dim posx&(1 To 1000), posy&(1 To 1000), angle!(1 To 1000)
Dim pi!, time!, k!
x = 0
Do
time = Timer
pi = 3.14159

ReDim pos(10000)
RandomizeX
For i = 1 To 1000
    posx(i) = 0
    posy(i) = 0
        For j = 1 To 1000 + x * 100
            pos(j) = 2 * pi * RndX
            posx(i) = posx(i) + Cos(pos(j))
            posy(i) = posy(i) + Sin(pos(j))
        Next j
        If posx(i) > 0 And posy(i) >= 0 Then
        angle(i) = Atn(posy(i) / posx(i)) * 180 / pi
        
        ElseIf posx(i) < 0 And posy(i) >= 0 Then
        angle(i) = Atn(posy(i) / posx(i)) * 180 / pi + 180
        
        ElseIf posx(i) < 0 And posy(i) < 0 Then
        angle(i) = Atn(posy(i) / posx(i)) * 180 / pi + 180
        
        ElseIf posx(i) > 0 And posy(i) < 0 Then
        angle(i) = Atn(posy(i) / posx(i)) * 180 / pi + 360
        
        ElseIf posx(i) = 0 And posy(i) > 0 Then
        angle(i) = 90
        
        ElseIf posx(i) = 0 And posy(i) < 0 Then
        angle(i) = 270
        
        End If
Next i

    sum_posx = 0
    sum_posy = 0
    sum_angle = 0
    sum_distance = 0
    
For i = 1 To 1000
        
sum_posx = sum_posx + posx(i)
sum_posy = sum_posy + posy(i)
sum_angle = sum_angle + angle(i)
sum_distance = sum_distance + (posx(i) ^ 2 + posy(i) ^ 2)
Next i
    
    avg_posx = sum_posx / 1000
    avg_posy = sum_posy / 1000
    avg_distance = Sqr(avg_posx ^ 2 + avg_posy ^ 2)
    avg_angle = sum_angle / 1000
    k = Sqr(sum_distance / 1000)
    
ActiveWorkbook.Worksheets("工作表2").Select
    ActiveSheet.Cells(3, 2).Value = "跳動次數"
    ActiveSheet.Cells(3, 3).Value = "平均距離"
    ActiveSheet.Cells(3, 4).Value = "平均方位角"
    ActiveSheet.Cells(3, 5).Value = "距離方均根"
    ActiveSheet.Cells(3, 6).Value = "執行時間"
       
    ActiveSheet.Cells(x + 4, 2) = 1000 + x * 100
    ActiveSheet.Cells(x + 4, 3) = avg_distance
    ActiveSheet.Cells(x + 4, 4) = avg_angle
    ActiveSheet.Cells(x + 4, 5) = k
    ActiveSheet.Cells(x + 4, 6) = Timer - time
    x = x + 1
    Loop Until x > 90
End Sub
```

---


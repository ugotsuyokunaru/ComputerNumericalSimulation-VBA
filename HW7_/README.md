# 這是個找質數的VBA實作
請打開excel >> 點選開發人員 >> 點選Visual Basic >> 觀看程式碼

---
以下為程式碼：

# 亂數模擬擲骰子
> 用亂數模擬擲骰子一千、一萬、十萬與一百萬次的結果，
> 將四項模擬的各15種情形出現的次數列表。

```VBA
Option Explicit
Sub 擲骰子()

Dim m!, n(0 To 3) As Single
Dim j!, i!, k!, dice(1 To 3) As Single
Dim t1!, t2!, t3!, t4!, t5!, t6!, t7!, t8!, t9!, t10!, t11!, t12!, t13!, t14!, t15!
ActiveSheet.Cells.Clear

RandomizeX
n(0) = 1000
n(1) = 10000
n(2) = 100000
n(3) = 1000000
    
For m = 0 To 3

Cells(4, 3 + 9 * m) = "骰子1"
Cells(4, 4 + 9 * m) = "骰子2"
Cells(4, 5 + 9 * m) = "骰子3"
Cells(4, 6 + 9 * m) = "情形"

Cells(4, 8 + 9 * m) = "情形列表"
Cells(4, 9 + 9 * m) = "次數列表"
Cells(4, 10 + 9 * m) = "機率列表"

Cells(5, 8 + 9 * m) = "123情形"
Cells(6, 8 + 9 * m) = "456情形"
Cells(7, 8 + 9 * m) = "1點"
Cells(8, 8 + 9 * m) = "2點"
Cells(9, 8 + 9 * m) = "3點"
Cells(10, 8 + 9 * m) = "4點"
Cells(11, 8 + 9 * m) = "5點"
Cells(12, 8 + 9 * m) = "6點"
Cells(13, 8 + 9 * m) = "1豹子"
Cells(14, 8 + 9 * m) = "2豹子"
Cells(15, 8 + 9 * m) = "3豹子"
Cells(16, 8 + 9 * m) = "4豹子"
Cells(17, 8 + 9 * m) = "5豹子"
Cells(18, 8 + 9 * m) = "6豹子"
Cells(19, 8 + 9 * m) = "沒有點數重擲"

t1 = 0
t2 = 0
t3 = 0
t4 = 0
t5 = 0
t6 = 0
t7 = 0
t8 = 0
t9 = 0
t10 = 0
t11 = 0
t12 = 0
t13 = 0
t14 = 0
t15 = 0

    For j = 1 To n(m)
    
        For i = 1 To 3
        k = RndX
        
            If 0 <= k And k < (1 / 6) Then
                    dice(i) = 1
                ElseIf 1 / 6 <= k And k < (2 / 6) Then
                    dice(i) = 2
                ElseIf 2 / 6 <= k And k < (3 / 6) Then
                    dice(i) = 3
                ElseIf 3 / 6 <= k And k < (4 / 6) Then
                    dice(i) = 4
                ElseIf 4 / 6 <= k And k < (5 / 6) Then
                    dice(i) = 5
                ElseIf 5 / 6 <= k And k < 1 Then
                    dice(i) = 6
            End If
            
        Next
                
        If (dice(1) = 1 And dice(2) = 2 And dice(3) = 3) Or (dice(1) = 1 And dice(3) = 2 And dice(2) = 3) Or (dice(2) = 1 And dice(1) = 2 And dice(3) = 3) Or (dice(2) = 1 And dice(3) = 2 And dice(1) = 3) Or (dice(3) = 1 And dice(2) = 2 And dice(1) = 3) Or (dice(3) = 1 And dice(1) = 2 And dice(2) = 3) Then
            Cells(4 + j, 3 + 9 * m) = dice(1)
            Cells(4 + j, 4 + 9 * m) = dice(2)
            Cells(4 + j, 5 + 9 * m) = dice(3)
            Cells(4 + j, 6 + 9 * m) = "123情形"
            t1 = t1 + 1
            
        ElseIf (dice(1) = 4 And dice(2) = 5 And dice(3) = 6) Or (dice(1) = 4 And dice(3) = 5 And dice(2) = 6) Or (dice(2) = 4 And dice(1) = 5 And dice(3) = 6) Or (dice(2) = 4 And dice(3) = 5 And dice(1) = 6) Or (dice(3) = 4 And dice(2) = 5 And dice(1) = 6) Or (dice(3) = 4 And dice(1) = 5 And dice(2) = 6) Then
            Cells(4 + j, 3 + 9 * m) = dice(1)
            Cells(4 + j, 4 + 9 * m) = dice(2)
            Cells(4 + j, 5 + 9 * m) = dice(3)
            Cells(4 + j, 6 + 9 * m) = "456情形"
            t2 = t2 + 1
        
        ElseIf (dice(1) = dice(2) And dice(2) = 1 And dice(2) <> dice(3)) Or (dice(2) = dice(3) And dice(3) = 1 And dice(3) <> dice(1)) Or (dice(1) = dice(3) And dice(3) = 1 And dice(3) <> dice(2)) Then
            Cells(4 + j, 3 + 9 * m) = dice(1)
            Cells(4 + j, 4 + 9 * m) = dice(2)
            Cells(4 + j, 5 + 9 * m) = dice(3)
            Cells(4 + j, 6 + 9 * m) = "1點"
            t3 = t3 + 1
        ElseIf (dice(1) = dice(2) And dice(2) = 2 And dice(2) <> dice(3)) Or (dice(2) = dice(3) And dice(3) = 2 And dice(3) <> dice(1)) Or (dice(1) = dice(3) And dice(3) = 2 And dice(3) <> dice(2)) Then
            Cells(4 + j, 3 + 9 * m) = dice(1)
            Cells(4 + j, 4 + 9 * m) = dice(2)
            Cells(4 + j, 5 + 9 * m) = dice(3)
            Cells(4 + j, 6 + 9 * m) = "2點"
            t4 = t4 + 1
        ElseIf (dice(1) = dice(2) And dice(2) = 3 And dice(2) <> dice(3)) Or (dice(2) = dice(3) And dice(3) = 3 And dice(3) <> dice(1)) Or (dice(1) = dice(3) And dice(3) = 3 And dice(3) <> dice(2)) Then
            Cells(4 + j, 3 + 9 * m) = dice(1)
            Cells(4 + j, 4 + 9 * m) = dice(2)
            Cells(4 + j, 5 + 9 * m) = dice(3)
            Cells(4 + j, 6 + 9 * m) = "3點"
            t5 = t5 + 1
        ElseIf (dice(1) = dice(2) And dice(2) = 4 And dice(2) <> dice(3)) Or (dice(2) = dice(3) And dice(3) = 4 And dice(3) <> dice(1)) Or (dice(1) = dice(3) And dice(3) = 4 And dice(3) <> dice(2)) Then
            Cells(4 + j, 3 + 9 * m) = dice(1)
            Cells(4 + j, 4 + 9 * m) = dice(2)
            Cells(4 + j, 5 + 9 * m) = dice(3)
            Cells(4 + j, 6 + 9 * m) = "4點"
            t6 = t6 + 1
        ElseIf (dice(1) = dice(2) And dice(2) = 5 And dice(2) <> dice(3)) Or (dice(2) = dice(3) And dice(3) = 5 And dice(3) <> dice(1)) Or (dice(1) = dice(3) And dice(3) = 5 And dice(3) <> dice(2)) Then
            Cells(4 + j, 3 + 9 * m) = dice(1)
            Cells(4 + j, 4 + 9 * m) = dice(2)
            Cells(4 + j, 5 + 9 * m) = dice(3)
            Cells(4 + j, 6 + 9 * m) = "5點"
            t7 = t7 + 1
        ElseIf (dice(1) = dice(2) And dice(2) = 6 And dice(2) <> dice(3)) Or (dice(2) = dice(3) And dice(3) = 6 And dice(3) <> dice(1)) Or (dice(1) = dice(3) And dice(3) = 6 And dice(3) <> dice(2)) Then
            Cells(4 + j, 3 + 9 * m) = dice(1)
            Cells(4 + j, 4 + 9 * m) = dice(2)
            Cells(4 + j, 5 + 9 * m) = dice(3)
            Cells(4 + j, 6 + 9 * m) = "6點"
            t8 = t8 + 1
            
        ElseIf dice(1) = dice(2) And dice(2) = dice(3) And dice(2) = 1 Then
            Cells(4 + j, 3 + 9 * m) = dice(1)
            Cells(4 + j, 4 + 9 * m) = dice(2)
            Cells(4 + j, 5 + 9 * m) = dice(3)
            Cells(4 + j, 6 + 9 * m) = "1豹子"
            t9 = t9 + 1
        ElseIf dice(1) = dice(2) And dice(2) = dice(3) And dice(2) = 2 Then
            Cells(4 + j, 3 + 9 * m) = dice(1)
            Cells(4 + j, 4 + 9 * m) = dice(2)
            Cells(4 + j, 5 + 9 * m) = dice(3)
            Cells(4 + j, 6 + 9 * m) = "2豹子"
            t10 = t10 + 1
        ElseIf dice(1) = dice(2) And dice(2) = dice(3) And dice(2) = 3 Then
            Cells(4 + j, 3 + 9 * m) = dice(1)
            Cells(4 + j, 4 + 9 * m) = dice(2)
            Cells(4 + j, 5 + 9 * m) = dice(3)
            Cells(4 + j, 6 + 9 * m) = "3豹子"
            t11 = t11 + 1
        ElseIf dice(1) = dice(2) And dice(2) = dice(3) And dice(2) = 4 Then
            Cells(4 + j, 3 + 9 * m) = dice(1)
            Cells(4 + j, 4 + 9 * m) = dice(2)
            Cells(4 + j, 5 + 9 * m) = dice(3)
            Cells(4 + j, 6 + 9 * m) = "4豹子"
            t12 = t12 + 1
        ElseIf dice(1) = dice(2) And dice(2) = dice(3) And dice(2) = 5 Then
            Cells(4 + j, 3 + 9 * m) = dice(1)
            Cells(4 + j, 4 + 9 * m) = dice(2)
            Cells(4 + j, 5 + 9 * m) = dice(3)
            Cells(4 + j, 6 + 9 * m) = "5豹子"
            t13 = t13 + 1
        ElseIf dice(1) = dice(2) And dice(2) = dice(3) And dice(2) = 6 Then
            Cells(4 + j, 3 + 9 * m) = dice(1)
            Cells(4 + j, 4 + 9 * m) = dice(2)
            Cells(4 + j, 5 + 9 * m) = dice(3)
            Cells(4 + j, 6 + 9 * m) = "6豹子"
            t14 = t14 + 1
            
        Else
            Cells(4 + j, 3 + 9 * m) = dice(1)
            Cells(4 + j, 4 + 9 * m) = dice(2)
            Cells(4 + j, 5 + 9 * m) = dice(3)
            Cells(4 + j, 6 + 9 * m) = "沒有點數重擲"
            t15 = t15 + 1
            
        End If
    Next
    
    Cells(5, 9 + 9 * m) = t1
    Cells(6, 9 + 9 * m) = t2
    Cells(7, 9 + 9 * m) = t3
    Cells(8, 9 + 9 * m) = t4
    Cells(9, 9 + 9 * m) = t5
    Cells(10, 9 + 9 * m) = t6
    Cells(11, 9 + 9 * m) = t7
    Cells(12, 9 + 9 * m) = t8
    Cells(13, 9 + 9 * m) = t9
    Cells(14, 9 + 9 * m) = t10
    Cells(15, 9 + 9 * m) = t11
    Cells(16, 9 + 9 * m) = t12
    Cells(17, 9 + 9 * m) = t13
    Cells(18, 9 + 9 * m) = t14
    Cells(19, 9 + 9 * m) = t15
    
    Cells(5, 10 + 9 * m) = t1 / n(m)
    Cells(6, 10 + 9 * m) = t2 / n(m)
    Cells(7, 10 + 9 * m) = t3 / n(m)
    Cells(8, 10 + 9 * m) = t4 / n(m)
    Cells(9, 10 + 9 * m) = t5 / n(m)
    Cells(10, 10 + 9 * m) = t6 / n(m)
    Cells(11, 10 + 9 * m) = t7 / n(m)
    Cells(12, 10 + 9 * m) = t8 / n(m)
    Cells(13, 10 + 9 * m) = t9 / n(m)
    Cells(14, 10 + 9 * m) = t10 / n(m)
    Cells(15, 10 + 9 * m) = t11 / n(m)
    Cells(16, 10 + 9 * m) = t12 / n(m)
    Cells(17, 10 + 9 * m) = t13 / n(m)
    Cells(18, 10 + 9 * m) = t14 / n(m)
    Cells(19, 10 + 9 * m) = t15 / n(m)

Next

End Sub
```

---

> 根據機率理論用 Excel 函數，計算問題二中15種情形的真實機率。
> (詳見 excel 檔案)

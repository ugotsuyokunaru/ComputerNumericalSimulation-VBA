# 這是個模擬棋子在圍棋盤上，隨機移動的VBA實作
請打開excel >> 點選開發人員 >> 點選Visual Basic >> 觀看程式碼與工作表上結果呈現

---
以下為程式碼：

# 亂數模擬棋子在圍棋盤( (-9,-9) 至 (9,9) )上隨機移動
> 計算至棋子離開棋盤為止  
> 1. 移動步數、  
> 2. 移出棋盤前位置、  
> 3. 移出棋盤後位置、  
> 4. 移出位置分布(藉此判斷邊或角)、  
> 5. 離開棋盤(x或y座標其中一個為10或-10)時，  
>     是從邊(非10的另一座標為-4~4)或角(非10的另一座標為>4或<-4)。  
>   
> 程式的邏輯：  
> 使用select case、Int函數與RndX函數亂數隨機選擇往上、下、左、右四個方向之一移動，  
> 且每移動一次計數一次，並用 do loop迴圈包以上過程，  
> loop until X座標或Y座標其中一個達到正負10即表示移動出圍棋盤，  
> 輸出移出後的座標位置與移動步數，再用多個if與elseif判定移動出去之前的座標位置，  
> 即結束第一次run。最外面再用for next迴圈包，到m=1000000次為止。  
>   

```VBA
Option Explicit
Sub 圍棋移動()

Dim m!, k!, j!, posx!(), posy!()
RandomizeX
ActiveSheet.Range("C6:H1000005").Clear
Cells(5, 3) = "執行次數"
Cells(5, 4) = "移動步數"
Cells(5, 5) = "移出棋盤前位置"
Cells(5, 6) = "移出棋盤後位置"
Cells(5, 7) = "移出位置分布"
Cells(5, 8) = "邊或角"


Cells(3, 4) = "平均步數"

For m = 1 To 1000000

ReDim posx(1 To m) As Single
ReDim posy(1 To m) As Single

posx(m) = 0
posy(m) = 0  '原點在棋盤中心
j = 0

    Do
        k = Int(RndX * 4) + 1
        
        Select Case k
            Case 1
                posx(m) = posx(m) + 1
            Case 2
                posx(m) = posx(m) - 1
            Case 3
                posy(m) = posy(m) + 1
            Case 4
                posy(m) = posy(m) - 1
        End Select
        
        j = j + 1
        
    Loop Until posx(m) = 10 Or posy(m) = 10 Or posx(m) = -10 Or posy(m) = -10
      
    
    Cells(m + 5, 3) = m
    Cells(m + 5, 4) = j
    
    
    If posx(m) = 10 Then
        Cells(m + 5, 5) = "( " & posx(m) - 1 & " , " & posy(m) & " )"
        Cells(m + 5, 7) = posy(m)
        If posy(m) <= 4 And posy(m) >= -4 Then
            Cells(m + 5, 8) = "邊"
        Else
            Cells(m + 5, 8) = "角"
        End If
        
    ElseIf posx(m) = -10 Then
        Cells(m + 5, 5) = "( " & posx(m) + 1 & " , " & posy(m) & " )"
        Cells(m + 5, 7) = posy(m)
        If posy(m) <= 4 And posy(m) >= -4 Then
            Cells(m + 5, 8) = "邊"
        Else
            Cells(m + 5, 8) = "角"
        End If

    ElseIf posy(m) = 10 Then
        Cells(m + 5, 5) = "( " & posx(m) & " , " & posy(m) - 1 & " )"
        Cells(m + 5, 7) = posx(m)
        If posx(m) <= 4 And posx(m) >= -4 Then
            Cells(m + 5, 8) = "邊"
        Else
            Cells(m + 5, 8) = "角"
        End If

    ElseIf posy(m) = -10 Then
        Cells(m + 5, 5) = "( " & posx(m) & " , " & posy(m) + 1 & " )"
        Cells(m + 5, 7) = posx(m)
        If posx(m) <= 4 And posx(m) >= -4 Then
            Cells(m + 5, 8) = "邊"
        Else
            Cells(m + 5, 8) = "角"
        End If

    End If
    
    Cells(m + 5, 6) = "( " & posx(m) & " , " & posy(m) & " )"

Next

End Sub
```

---

> 繪出移出次數對移出位置的分佈圖。  
> (詳見 excel 檔案)
>   

移出位置分布 |	次數
------------ | -------------
-9 |	10858
-8 |	21663
-7 |	32303
-6 |	43230
-5 |	53308
-4 |	63031
-3 |	71806
-2 |	78495
-1 |	82802
0 |	84055
1 |	82824
2 |	78448
3 |	71651
4 |	63081
5 |	53870
6 |	43555
7 |	32613
8 |	21739
9 |	10668

>   
> 分佈圖大致呈現較平滑的常態分佈。  

> 觀察發現，從邊移出去的棋子會是從角落移出去的2倍。  
> 因為從邊出去的一定會比從角出去的多，角距離圓心較遠。

---



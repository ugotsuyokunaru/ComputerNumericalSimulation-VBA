# 這是__________的VBA實作
請打開excel >> 點選開發人員 >> 點選Visual Basic >> 觀看程式碼與工作表上結果呈現

---

以下為程式碼：

```VBA
Option Explicit
Sub 五十乘五十stack10轉12版()

ActiveSheet.Cells.Clear
        
Dim water%, x%, i%, j%, stacking%, a%, hole%, k As Integer
Dim n1!, n2!, n3 As Single
Dim stack(1 To 2500) As Single
Dim stTimer As Single
Dim pos As Range
RandomizeX
'Range("C3:N14").Interior.ColorIndex = 16
'Range("D3:M3").Interior.ColorIndex = 28

ActiveSheet.Cells(2, "BD") = "模擬序數"
ActiveSheet.Cells(2, "BE") = "移除格子點數"
ActiveSheet.Cells(2, "BF") = "切斷孔隙率"
ActiveSheet.Cells(2, "BG") = "跑程式時間"


Set pos = Range("C3:BB54")

'For a = 1 To 2500
'stack(a) = 0
'Next

Range("D4:BA53").Select
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
    
'----------前置完成------------------------------------------------------

For k = 1 To 100

stTimer = Timer

    pos.Cells().Interior.ColorIndex = 16
    For i = 2 To 51
    pos.Cells(i).Interior.ColorIndex = 28
    Next
    
    Do  '先暫時用for迴圈代替
    'For a = 1 To 100
        
        '挖洞
        x = Int(RndX * 2500) + 1
        If pos(x + 53 + (Int((x - 1) / 50)) * 2).Interior.ColorIndex = 16 Then
            pos(x + 53 + (Int((x - 1) / 50)) * 2).Interior.ColorIndex = xlNone '10x10的個數"x"，加上最上排n+2個、次排第一個(共n+3個)，再加int(x/n)*2個
                   
            '挖完洞，判斷流水染色，如果四周(左右+1/上下+(n+2))有水，才做
            If pos(x + 53 + (Int((x - 1) / 50)) * 2 + 1).Interior.ColorIndex = 28 Or pos(x + 53 + (Int((x - 1) / 50)) * 2 - 1).Interior.ColorIndex = 28 _
            Or pos(x + 53 + (Int((x - 1) / 50)) * 2 + 52).Interior.ColorIndex = 28 Or pos(x + 53 + (Int((x - 1) / 50)) * 2 - 52).Interior.ColorIndex = 28 Then
        
                '每染到一個新白格時，先記錄新白格周圍能繼續染的白格到stack中暫存
                Do
                
                j = 1
                pos(x + 53 + (Int((x - 1) / 50)) * 2).Interior.ColorIndex = 28   '新挖的洞本身先染水
                
                '左
                If pos(x + 53 + (Int((x - 1) / 50)) * 2 - 1).Interior.ColorIndex = xlNone Then
                j = j + 1
                stack(j) = x + 53 + (Int((x - 1) / 50)) * 2 - 1
                End If
                
                '右
                If pos(x + 53 + (Int((x - 1) / 50)) * 2 + 1).Interior.ColorIndex = xlNone Then
                j = j + 1
                stack(j) = x + 53 + (Int((x - 1) / 50)) * 2 + 1
                End If
                
                '上
                If pos(x + 53 + (Int((x - 1) / 50)) * 2 - 52).Interior.ColorIndex = xlNone Then
                j = j + 1
                stack(j) = x + 53 + (Int((x - 1) / 50)) * 2 - 52
                End If
                
                '下
                If pos(x + 53 + (Int((x - 1) / 50)) * 2 + 52).Interior.ColorIndex = xlNone Then
                j = j + 1
                stack(j) = x + 53 + (Int((x - 1) / 50)) * 2 + 52
                End If
                
                    '執行暫存中的stack(一個一個 j 依序執行)，後進先出
                    If j > 1 Then
                        Do
                            pos(stack(j)).Interior.ColorIndex = 28  '把下一個洞染水
                            stacking = stack(j)                     '存出新洞位置暫存到stacking(以免stack(j)撞到)
                            j = j - 1   '染完且存完新洞之後j=j-1
                            
                            '先存入stack
                            '左
                            If pos(stacking - 1).Interior.ColorIndex = xlNone Then
                            j = j + 1
                            stack(j) = stacking - 1
                            End If
                    
                            '右
                            If pos(stacking + 1).Interior.ColorIndex = xlNone Then
                            j = j + 1
                            stack(j) = stacking + 1
                            End If
                    
                            '上
                            If pos(stacking - 52).Interior.ColorIndex = xlNone Then
                            j = j + 1
                            stack(j) = stacking - 52
                            End If
                    
                            '下
                            If pos(stacking + 52).Interior.ColorIndex = xlNone Then
                            j = j + 1
                            stack(j) = stacking + 52
                            End If
                            
                            '做完其中一個方向(固定左右下上的順序)再回stack(j)做下一條路，做到
                        Loop Until j = 1
                    End If
                                
                Loop Until j = 1
        
            End If
            
        
                'stack全部流完，還沒碰到底，才做下一輪
                
                water = 0
                For i = 2602 To 2651
                    If pos(i).Interior.ColorIndex = 28 Then   'pos(122~131)其中任一通水就結束   n*(n-2)+2 ~ n*(n-1)-1
                    water = 1
                    End If
                Next
            
                If water = 1 Then
                    pos(x + 53 + (Int((x - 1) / 50)) * 2).Interior.ColorIndex = 7
                End If
            
        End If
    
    'Next
    Loop Until water = 1
    
    hole = 0
    
    For i = 1 To 2704
        If pos(i).Interior.ColorIndex = xlNone Or pos(i).Interior.ColorIndex = 28 Then
        hole = hole + 1
        End If
    Next

    ActiveSheet.Cells(2 + k, "BD") = k
    ActiveSheet.Cells(2 + k, "BE") = hole
    ActiveSheet.Cells(2 + k, "BF") = hole / 2500
    ActiveSheet.Cells(2 + k, "BG") = Timer - stTimer

Next

n1 = 0
n2 = 0
n3 = 0

For i = 1 To 100
n1 = n1 + ActiveSheet.Cells(2 + i, "BE").Value
n2 = n2 + ActiveSheet.Cells(2 + i, "BF").Value
n3 = n3 + ActiveSheet.Cells(2 + i, "BG").Value
Next

ActiveSheet.Cells(103, "BC") = "平均"
ActiveSheet.Cells(103, "BE") = n1 / 100
ActiveSheet.Cells(103, "BF") = n2 / 100
ActiveSheet.Cells(103, "BG") = n3 / 100

' 調整格式 巨集
    Range("BE103").Select
    Selection.NumberFormatLocal = "0_ "
    ActiveWindow.SmallScroll Down:=-102
    Range("BF3").Select
    Range(Selection, Selection.End(xlDown)).Select
    Selection.NumberFormatLocal = "0.00%"
    ActiveWindow.SmallScroll Down:=48

```

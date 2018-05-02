# 這是個找質數的VBA實作
請打開excel >> 點選開發人員 >> 點選Visual Basic >> 觀看程式碼

---
以下為程式碼
---

>Option Explicit  
>Sub 找質數()  
>Dim i%, j%, p%, w As Integer  
>Dim m As Long  
>  
>ActiveSheet.Cells.Clear  
>m = InputBox("請輸入您希望質數找到哪個數字？")  
>p = 2  
>  
>For i = 2 To m  
>    \\\w = 0\\\  
>      
>    \For j = 2 To (i - 1)  
>        \If i Mod j = 0 Then  
>        \w = 1  
>        \End If  
>    \Next j  
>      
>    \If w = 0 Then  
>        \ActiveSheet.Cells(p, 2).Value = i  
>        \p = p + 1  
>    \End If  

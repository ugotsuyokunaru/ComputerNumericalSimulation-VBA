# 這是個取亂數的程式碼，
# 可以給所有需要隨機取亂數的VBA程式套用
請打開excel >> 點選開發人員 >> 點選Visual Basic >> 觀看程式碼

---
以下為程式碼：

> 呼叫時，  
> 原本VBA中的Randomize在此應以RandomizeX取代  
> 原本VBA中的Rnd在此應以RndX取代  
>   

```VBA
'從以下網址抓來  http://www.vbforums.com/showthread.php?499661-Wichmann-Hill-Pseudo-Random-Number-Generator-an-alternative-for-VB-Rnd()-function
'長週期亂數產生器程式碼，週期有10兆。
'凡是亂數的數目超過100萬，都應該使用此程式碼
'使用時，將此程式碼複製到在VBE視窗之Excel工作表的一個新模組內即可
'原本VBA中的Randomize在此應以RandomizeX取代
'原本VBA中的Rnd在此應以RndX取代

'============================================================================
'http://support.microsoft.com/kb/828795
'
'The basic idea is that if you take three random numbers on [0,1] and sum them,
'the fractional part of the sum is itself a random number on [0,1].
'The critical statements in the Fortran code listing from the original
'Wichman and Hill article are:
'
'C  IX, IY, IZ SHOULD BE SET TO INTEGER VALUES BETWEEN 1 AND 30000 BEFORE FIRST ENTRY
'IX = MOD(171 * IX, 30269)
'IY = MOD(172 * IY, 30307)
'IZ = MOD(170 * IZ, 30323)
'RANDOM = AMOD(FLOAT(IX) / 30269.0 + FLOAT(IY) / 30307.0 + FLOAT(IZ) / 30323.0, 1.0)
'=======================================================================

Option Explicit
Private ix As Long, iy As Long, iz As Long

Sub RandomizeX(Optional ByVal Number)
   Const MaxLong As Double = 2 ^ 31 - 1
   Dim n As Long
   Dim d As Double
   
   If IsMissing(Number) Then
      n = Timer * 60
      '-- Timer is only updated in every 1/60th second.
      '-- Multiply by 60 to reduce the chance of seed
      '-- to be repeated in subsequence calls of RandomizedX
   Else
      d = Abs(Int(Val(Number)))
      If d > MaxLong Then '-- prevent Long overflow
         d = d - Int(d / MaxLong) * MaxLong
      End If
      n = d
   End If
   ix = (n Mod 30269)
   iy = (n Mod 30307)
   iz = (n Mod 30323)
   '-- ix, iy, iz cannot be 0
   If ix = 0 Then ix = 171
   If iy = 0 Then ix = 172
   If iz = 0 Then ix = 170
End Sub

Function RndX(Optional ByVal Number As Long = 1) As Double
   Dim r As Double
   
   If ix = 0 Then '-- ix, iy, iz cannot be 0.
      ix = 171    '-- Initial values of ix, iy and iz are 0.
      iy = 172    '-- If any of these equal 0,
      iz = 170    '-- it will be stucked with 0 forever.
   End If
   If Number <> 0 Then
      If Number < 0 Then
         RandomizeX Number
      End If
      ix = (171 * ix) Mod 30269 '-- This has been tested:
      iy = (172 * iy) Mod 30307 '-- ix, iy, iz will never be 0.
      iz = (170 * iz) Mod 30323 '-- ---------------------------
   End If
   r = ix / 30269# + iy / 30307# + iz / 30323#
   RndX = r - Int(r)
End Function
```

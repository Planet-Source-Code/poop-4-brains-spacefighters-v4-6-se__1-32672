Attribute VB_Name = "modGrad"
Enum Direction
dHorizontal = 0
dVertical = 1
End Enum

Function Gradient(board As Object, Dir As Direction, r1, g1, b1, r2, g2, b2)
On Error Resume Next
Dim i
Dim r, g, b, max

Dim rc, gc, bc

rc = r2 - r1
bc = b2 - b1
gc = g2 - g1

MakePos rc
MakePos bc
MakePos gc

Select Case Dir
Case Direction.dHorizontal
gc = board.ScaleWidth \ gc
bc = board.ScaleWidth \ bc
rc = board.ScaleWidth \ rc
max = board.ScaleWidth
Case Direction.dVertical
gc = board.ScaleHeight \ gc
bc = board.ScaleHeight \ bc
rc = board.ScaleHeight \ rc
max = board.ScaleHeight
End Select

r = r1
g = g1
b = b1

For i = 0 To max
Select Case Dir
Case Direction.dHorizontal
board.Line (i, 0)-(i, board.ScaleHeight), RGB(r, g, b)
Case Direction.dVertical
board.Line (0, i)-(board.ScaleWidth, i), RGB(r, g, b)
End Select

If r < r2 Then r = r + rc
If r > r2 Then r = r - rc

If b < b2 Then b = b + bc
If b > b2 Then b = b - bc

If g < g2 Then g = g + gc
If g > g2 Then g = g - gc
Next i
End Function

Function MakePos(number)
If number < 0 Then number = -number
End Function


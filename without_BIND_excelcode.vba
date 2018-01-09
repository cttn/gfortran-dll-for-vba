Option Base 1
Option Explicit

Declare PtrSafe Function exarrays Lib "D:\exVBA.dll" Alias "exarrays@8" (ByVal n As Long, x As Double) As Double
Declare Sub subarrays Lib "D:\exVBA.dll" Alias "subarrays@12" (ByVal n As Long, x As Double, res As Double)

Public Function foo_array() As Double
  Dim i As Long
  Dim dval(5) As Double
  For i = 1 To 5
    dval(i) = CDbl(i)
  Next
  foo_array = exarrays(5, dval(1))
End Function

Public Function gen_array(x As Range)
  Dim n As Long
  n = x.Count
  Dim a() As Double
  ReDim a(n) As Double
  a = vector2array(x)
  Dim res() As Double
  ReDim res(n) As Double
  
  Call subarrays(n, a(1), res(1))
  gen_array = res
End Function

Private Function vector2array(x As Range) As Double()
  Dim n As Long
  n = x.Count
  Dim a() As Double
  ReDim a(n) As Double
  Dim i As Long
  For i = 1 To n
    a(i) = CDbl(x.Item(i))  'Cast to double
  Next
  vector2array = a
End Function

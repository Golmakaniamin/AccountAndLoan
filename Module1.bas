Attribute VB_Name = "Amin"
Dim strd As String, d1 As String, m1 As String, y1 As String, d2 As String, m2 As String, y2 As String, strdate1 As String
Dim yt As String, mt As String, dt As String, sum As String

Public Function dateaminEktelaf(date1 As String, date2 As String) As String
If date1 <= date2 Then
  y1 = Mid(date1, 1, 4)
  m1 = Mid(date1, 6, 2)
  d1 = Mid(date1, 9, 2)

  y2 = Mid(date2, 1, 4)
  m2 = Mid(date2, 6, 2)
  d2 = Mid(date2, 9, 2)

  yt = y2 - y1
  If yt = 0 Then
    mt = m2 - m1
    If mt = 0 Then
      dateaminEktelaf = d2 - d1
    Else
      dateaminEktelaf = ((mt - 1) * 30) + (d2 + (30 - d1))
    End If
  Else
    dateaminEktelaf = ((yt - 1) * 365) + (((m2 + (12 - m1)) - 1) * 30) + (d2 + (30 - d1))
  End If
Else
dateaminEktelaf = "a"
End If
End Function

Public Function dateaminEktelafmoon(date1 As String, date2 As String) As String
If date1 <= date2 Then
  y1 = Mid(date1, 1, 4)
  m1 = Mid(date1, 6, 2)

  y2 = Mid(date2, 1, 4)
  m2 = Mid(date2, 6, 2)

  yt = y2 - y1
  If yt = 0 Then
    mt = m2 - m1
    If mt <> 0 Then dateaminEktelafmoon = mt
  Else
    dateaminEktelafmoon = ((yt - 1) * 12) + ((m2 + (12 - m1)))
  End If
Else
  dateaminEktelafmoon = "a"
End If
End Function

Public Function dateaminEzafeMoon(date1 As String, number As String) As String
y1 = Mid(date1, 1, 4)
m1 = Mid(date1, 6, 2)
d1 = Mid(date1, 9, 2)

m1 = Val(m1) + number
If m1 > 12 Then
  y1 = Val(y1) + Val(m1 \ 12)
  m1 = m1 Mod 12
  If m1 = 0 Then
    y1 = y1 - 1
    m1 = 12
  End If
End If
If Len(m1) = 1 Then m1 = "0" + m1
dateaminEzafeMoon = y1 + "/" + m1 + "/" + d1
End Function

Public Function dateaminEzafeday(date1 As String, number As String) As String
y1 = Mid(date1, 1, 4)
m1 = Mid(date1, 6, 2)
d1 = Mid(date1, 9, 2)

d1 = Val(d1) + number

If d1 > 30 Then
  m1 = Val(m1) + Val(d1 \ 30)
  d1 = d1 Mod 30
End If

If m1 > 12 Then
  y1 = Val(y1) + Val(m1 \ 12)
  m1 = m1 Mod 12
  If m1 = 0 Then
    y1 = y1 - 1
    m1 = 12
  End If
End If

If Len(d1) = 1 Then d1 = "0" + d1
If Len(m1) = 1 Then m1 = "0" + m1
dateaminEzafeday = y1 + "/" + m1 + "/" + d1
End Function

Public Function moneyaminjoda(number As String) As String
Dim q As String
Dim w As Integer
q = ""
For w = 1 To Len(number)
  If Mid(number, w, 1) <> "." Then
    q = q + Mid(number, w, 1)
  End If
Next w
number = q
q = ""
e = Len(number) Mod 3

For w = Len(number) + 1 To 1 Step -3
  q = Mid(number, w, 3) + "." + q
Next w

q = Left(number, e) + "." + q
If Left(q, 1) = "." Then q = Right(q, Len(q) - 1)
If Right(q, 2) = ".." Then q = Left(q, Len(q) - 2)
moneyaminjoda = q
End Function

Public Function moneyaminnojoda(number As String) As Long
Dim q As String
q = ""
For w = 1 To Len(number)
  If Mid(number, w, 1) <> "." Then
    q = q + Mid(number, w, 1)
  End If
Next w
moneyaminnojoda = q
End Function


Attribute VB_Name = "Module1"
Type sýra
x As Double
y As Double
End Type




Sub vikor()
x = InputBox("Alternatif Sayýsý Giriniz")
y = InputBox("Kriter Sayýsý Giriniz")
Dim etop() As Double
ReDim etop(y) As Double

Dim renk() As Double
ReDim renk(x, y) As Double
Dim w() As Double
ReDim w(y) As Double

Dim ext()
Dim AA() As Double
Dim BB() As Double
ReDim AA(y) As Double
ReDim BB(y) As Double
Dim R As Long
Dim c As Long
ReDim ext(y)
Dim fp() As Double
ReDim fp(y) As Double

Dim fn() As Double
ReDim fn(y) As Double

Dim S() As Double
ReDim S(x) As Double









For c = 2 To y + 1
R = 2
AA(c - 1) = Cells(R, c)
BB(c - 1) = Cells(R, c)
For R = 2 To x + 1
If Cells(R, c) > AA(c - 1) Then
AA(c - 1) = Cells(R, c)
End If
If Cells(R, c) < BB(c - 1) Then
BB(c - 1) = Cells(R, c)
End If
Next R
Next c


For c = 2 To y + 1
For R = 2 To x + 1
If Cells(R, c) < 0 Then
Cells(R, c) = (AA(c - 1) - BB(c - 1)) * (Cells(R, c) - Int(BB(c - 1))) / ((Sgn(AA(c - 1)) * Int(Abs(AA(c - 1)))) - (Sgn(BB(c - 1)) * Int(Abs(BB(c - 1)))))

End If

Next R
Next c
For c = 2 To y + 1
For R = 2 To x + 1
If Cells(R, c) = 0 Then
Cells(R, c) = 0.000001
End If
Next R
Next c



R = 2
 c = 1
For R = 2 To x + 1

Cells(R, c).Select
Cells(R, c).Value = "A" & CStr(R - 1)

Next R
R = 1


For c = 2 To y + 1

Cells(R, c).Select
Cells(R, c).Value = "C" & CStr(c - 1)
geri_:
ext(c - 1) = InputBox("Bu kriter için ideal deðer minimum ise 1 maximum ise 2 giriniz.")
If ext(c - 1) < 1 Or ext(c - 1) > 2 Then
MsgBox "Yanlýþ deðer girdiniz"
GoTo geri_

End If
Next c

Cells(x + 2, 1).Value = "W"

For c = 2 To y + 1
etop(c - 1) = 0

For R = 2 To x + 1
n = Cells(R, c)


etop(c - 1) = etop(c - 1) + n
Next R



Next c

For c = 2 To y + 1

For R = 2 To x + 1

renk(R - 1, c - 1) = Cells(R, c) / etop(c - 1)




Next R


Next c

P = 0
For c = 2 To y + 1
t = 0
For R = 2 To x + 1


k = -renk(R - 1, c - 1) * Log(renk(R - 1, c - 1)) / Log(x)
t = t + k
Next R
P = P + 1 - t
w(c - 1) = t



Next c



For c = 2 To y + 1
w(c - 1) = (1 - w(c - 1)) / P

Cells(x + 2, c).Value = w(c - 1)

Next c
Cells(x + 2, 1).Select
MsgBox "W ile baþlayan satýrda herbir kriter için Entropy yöntemiyle hesaplanmýþ olan aðýrlýk deðerleri yer almaktadýr"

For j = 1 To y

If ext(j) = 2 Then
fp(j) = AA(j)
fn(j) = BB(j)
Else
fp(j) = BB(j)
fn(j) = AA(j)
End If
Next j


'Sj deðerlerinin hesaplanmasý

'Rj deðerlerinin hesaplanmasý
Dim Rj() As Double
ReDim Rj(x) As Double

Cells(1, y + 4) = "S"
Cells(1, y + 5) = "R"

fa = -1E+22
na = 1E+28
Ry = 1E+31
Rn = -1E+34
For i = 1 To x

ma = 0
ka = -1E+29
For j = 1 To y

S(i) = w(j) * (fp(j) - Cells(i + 1, j + 1)) / (fp(j) - fn(j))
Rj(i) = S(i)

If ka < Rj(i) Then
ka = Rj(i)
End If

ma = ma + S(i)
Next j







Rj(i) = ka
S(i) = ma

'S*=min Sj Sy ile gösteriliyor
'S-=max Sj Sn ile gösteriliyor
If fa < S(i) Then
fa = S(i)
End If
If na > S(i) Then
na = S(i)
End If
'R*=min Rj Ry ile gösteriliyor
'R-=max Rj Rn ile gösteriliyor

If Ry > Rj(i) Then
Ry = Rj(i)
End If
If Rn < Rj(i) Then
Rn = Rj(i)
End If

Cells(i + 1, y + 4) = S(i)
Cells(i + 1, y + 5) = Rj(i)



Next i
Sn = fa
Sy = na
Cells(1, y + 6) = "S*"
Cells(2, y + 6) = Sy
Cells(1, y + 7) = "S-"
Cells(2, y + 7) = Sn
Cells(1, y + 8) = "R*"
Cells(2, y + 8) = Ry

Cells(1, y + 9) = "R-"
Cells(2, y + 9) = Rn
'Qj deðerlerinin hesaplanmasý
Cells(1, y + 10) = "Q"

Cells(1, y + 11) = "Q Sýra"
Cells(1, y + 12) = "Q Deðer"
Cells(1, y + 13) = "S Sýra"
Cells(1, y + 14) = "S Deðer"

Cells(1, y + 15) = "R Sýra"
Cells(1, y + 16) = "R Deðer"

Dim Qj() As Double
ReDim Qj(x) As Double
For j = 1 To x
Qj(j) = 0.5 * (S(j) - Sy) / (Sn - Sy) + 0.5 * (Rj(j) - Ry) / (Rn - Ry)
Cells(1 + j, y + 10) = Qj(j)
Next j
'Qj lerin sýralanmasý
'Sj lerin sýralanmasý
'Rj lerin sýralanmasý
Dim Q() As sýra
ReDim Q(x) As sýra
Dim Si() As sýra
ReDim Si(x) As sýra
Dim Ri() As sýra
ReDim Ri(x) As sýra

DQ = 1 / (x - 1)
For i = 1 To x
mes = Qj(i)
nes = S(i)
les = Rj(i)
For j = 1 To x
If les >= Rj(j) Then
les = Rj(j)
Ri(i).y = j
End If


If nes >= S(j) Then
nes = S(j)
Si(i).y = j
End If

If mes >= Qj(j) Then
mes = Qj(j)
Q(i).y = j
End If
Next j
Ri(i).x = les
Rj(Ri(i).y) = 1E+26


Si(i).x = nes
S(Si(i).y) = 1E+26
Q(i).x = mes
Qj(Q(i).y) = 1E+26
Cells(1 + i, y + 11) = "A" & Str(Q(i).y)
Cells(1 + i, y + 12) = Q(i).x
Cells(1 + i, y + 13) = "A" & Str(Si(i).y)
Cells(1 + i, y + 14) = Si(i).x
Cells(1 + i, y + 15) = "A" & Str(Ri(i).y)
Cells(1 + i, y + 16) = Ri(i).x


Next i

Cells(1, y + 17) = "Uzlaþýk çözüm kümesi"
Dim kosul(2) As Boolean


If (Q(2).x - Q(1).x) >= DQ Then
kosul(1) = True
Else
kosul(1) = False
End If
If Si(1).y = Q(1).y Or Ri(1).y = Q(1).y Then
kosul(2) = True
Else
kosul(2) = False
End If
If kosul(1) = False Then
For j = 3 To x

If (Q(j).x - Q(1).x) >= DQ Then

Exit For
End If
Next j

For i = 1 To j
Cells(i + 1, y + 18) = "A" & Str(Q(i).y)
Next i

End If
If kosul(1) = False Then
Cells(2, y + 18) = "A" & Str(Q(1).y)

Cells(3, y + 18) = "A" & Str(Q(2).y)
End If
If kosul(1) = True And kosul(2) = True Then
Cells(2, y + 18) = "A" & Str(Q(1).y)
End If

End Sub


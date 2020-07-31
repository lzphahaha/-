Private Sub Form_Load()

Dim a As Single
Dim q As Single
Dim k As Single
Dim v As Single
Dim d As Single



Combo1.AddItem "PE32"
Combo1.AddItem "PE40"
Combo1.AddItem "PE50"
Combo1.AddItem "PE63"
Combo1.AddItem "PE90"
Combo1.AddItem "PE110"
Combo1.AddItem "PE160"
Combo1.Text = "管径"
End Sub



Private Sub Command1_Click()

a = Val(Text1.Text)

Select Case a
Case 1
k = 1
Case 2
k = 0.56
Case 3
k = 0.44
Case 4
k = 0.38
Case 5
k = 0.68
Case 6
k = 0.31
Case 7
k = 0.29
Case 8
k = 0.27
Case 9
k = 0.26
Case 10
k = 0.25
Case 10 To 15
k = 0.22
Case 15 To 20
k = 0.21
Case 20 To 25
k = 0.2
Case 25 To 30
k = 0.19
Case 30 To 40
k = 0.18
Case 40 To 50
k = 0.178
Case 50 To 60
k = 0.176
Case 60 To 70
k = 0.174
Case 70 To 80
k = 0.172
Case 80 To 90
k = 0.171
Case 90 To 100
k = 0.17
Case 100 To 200
k = 0.16
Case 200 To 300
k = 0.15
Case 300 To 400
k = 0.14
Case 400 To 500
k = 0.138
Case 500 To 700
k = 0.134
Case 700 To 1000
k = 0.13
End Select

q = 2.5 * k * a

If Combo1.Text = "PE32" Then
d = 26
ElseIf Combo1.Text = "PE40" Then
d = 32.6
ElseIf Combo1.Text = "PE50" Then
d = 40.8
ElseIf Combo1.Text = "PE63" Then
d = 51.4
ElseIf Combo1.Text = "PE90" Then
d = 73.6
ElseIf Combo1.Text = "PE110" Then
d = 90
Else
d = 130.8
End If
Text5.Text = d

v = 4 * q / (0.36 * 3.14 * d * d)

Text2.Text = "燃气密度=0.82kg/m3" & vbCrLf & "运动粘度=12.56㎡/s" & vbCrLf & "计算流量=" & q & "m3/h"
Text3.Text = v

Dim r As Integer
Dim moca As Single
Dim yasun As Single
Dim l As Integer
l = Val(Text6.Text)

re = d * v / 12.56
If re <= 2100 Then
moca = 64 / re
yasun = 1.13 * 10000000000# * q * 12.56 * 0.82 * 1.073 * l / d / d / d / d
ElseIf 2100 < re <= 3500 Then
moca = 0.03 + (re - 2100) / (65 * re - 100000)
yasun = 1900000 * q * q * 0.82 * 1.073 * (1 + (11.8 * q - 70000 * d * 12.56) / (23 * q - 100000 * d * 12.56)) * l / d / d / d / d / d
Else
moca = 0.11 * (0.01 / d + 68 / re) ^ 0.25
yasun = 6900000 * q * q * 0.82 * 1.073 * l * (0.01 / d + 192.2 * d * 12.56) ^ 0.25 / d / d / d / d / d
End If

Text4.Text = yasun

End Sub

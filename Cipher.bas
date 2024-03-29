Attribute VB_Name = "Cipher"
' =================================================================
' Password Guard source code
' Version 1.5
' Copyright (C) 2000-2001 Khaery Rida
' =================================================================

' The PC1 Encryption Algorithm
' Cipher Strength: 128-bit
' Original version was written by Alexander Pukall
' alexandermail@hotmail.com
' =================================================================


Dim x1a0(9) As Long
Dim cle(17) As Long
Dim x1a2 As Long

Dim inter As Long, res As Long, ax As Long, bx As Long
Dim cx As Long, dx As Long, si As Long, tmp As Long
Dim i As Long, c As Byte


Sub Assemble()

x1a0(0) = ((cle(1) * 256) + cle(2)) Mod 65536
code
inter = res

x1a0(1) = x1a0(0) Xor ((cle(3) * 256) + cle(4))
code
inter = inter Xor res


x1a0(2) = x1a0(1) Xor ((cle(5) * 256) + cle(6))
code
inter = inter Xor res

x1a0(3) = x1a0(2) Xor ((cle(7) * 256) + cle(8))
code
inter = inter Xor res

x1a0(4) = x1a0(3) Xor ((cle(9) * 256) + cle(10))
code
inter = inter Xor res

x1a0(5) = x1a0(4) Xor ((cle(11) * 256) + cle(12))
code
inter = inter Xor res

x1a0(6) = x1a0(5) Xor ((cle(13) * 256) + cle(14))
code
inter = inter Xor res

x1a0(7) = x1a0(6) Xor ((cle(15) * 256) + cle(16))
code
inter = inter Xor res

i = 0

End Sub

Sub code()
dx = (x1a2 + i) Mod 65536
ax = x1a0(i)
cx = &H15A
bx = &H4E35

tmp = ax
ax = si
si = tmp

tmp = ax
ax = dx
dx = tmp

If (ax <> 0) Then
ax = (ax * bx) Mod 65536
End If

tmp = ax
ax = cx
cx = tmp

If (ax <> 0) Then
ax = (ax * si) Mod 65536
cx = (ax + cx) Mod 65536
End If

tmp = ax
ax = si
si = tmp
ax = (ax * bx) Mod 65536
dx = (cx + dx) Mod 65536

ax = ax + 1

x1a2 = dx
x1a0(i) = ax

res = ax Xor dx
i = i + 1

End Sub

Public Function crypt(ByVal inp As String, ByVal Key As String) As String

crypt = ""
si = 0
x1a2 = 0
i = 0

For fois = 1 To 16
cle(fois) = 0
Next fois

champ1 = Key
lngchamp1 = Len(champ1)

For fois = 1 To lngchamp1
cle(fois) = Asc(Mid(champ1, fois, 1))
Next fois

champ1 = inp
lngchamp1 = Len(champ1)
For fois = 1 To lngchamp1
c = Asc(Mid(champ1, fois, 1))

Assemble

cfc = (((inter / 256) * 256) - (inter Mod 256)) / 256
cfd = inter Mod 256

For compte = 1 To 16

cle(compte) = cle(compte) Xor c

Next compte

c = c Xor (cfc Xor cfd)

d = (((c / 16) * 16) - (c Mod 16)) / 16
e = c Mod 16

crypt = crypt + Chr$(&H61 + d) ' d+&h61 give one letter range from a to p for the 4 high bits of c
crypt = crypt + Chr$(&H61 + e) ' e+&h61 give one letter range from a to p for the 4 low bits of c


Next fois

End Function

Public Function decrypt(ByVal inp As String, ByVal Key As String) As String

decrypt = ""
si = 0
x1a2 = 0
i = 0

For fois = 1 To 16
cle(fois) = 0
Next fois

champ1 = Key
lngchamp1 = Len(champ1)

For fois = 1 To lngchamp1
cle(fois) = Asc(Mid(champ1, fois, 1))
Next fois

champ1 = inp
lngchamp1 = Len(champ1)

For fois = 1 To lngchamp1

d = Asc(Mid(champ1, fois, 1))
If (d - &H61) >= 0 Then
d = d - &H61  ' to transform the letter to the 4 high bits of c
If (d >= 0) And (d <= 15) Then
d = d * 16
End If
End If
If (fois <> lngchamp1) Then
fois = fois + 1
End If
e = Asc(Mid(champ1, fois, 1))
If (e - &H61) >= 0 Then
e = e - &H61 ' to transform the letter to the 4 low bits of c
If (e >= 0) And (e <= 15) Then
c = d + e
End If
End If

Assemble

cfc = (((inter / 256) * 256) - (inter Mod 256)) / 256
cfd = inter Mod 256

c = c Xor (cfc Xor cfd)

For compte = 1 To 16

cle(compte) = cle(compte) Xor c

Next compte

decrypt = decrypt + Chr$(c)

Next fois

End Function

Public Function cryptPassword(ByVal inp As String, ByVal Key As String) As PasswordEncodingFlags

Dim tmpAsc As Long
tmpAsc = 0
For currentCounter = 1 To Len(Key)
    tmpAsc = tmpAsc + Asc(Mid$(Key, currentCounter, 1))
Next

cryptPassword.encodedPassword = ""
si = 0
x1a2 = 0
i = 0

For fois = 1 To 16
    cle(fois) = 0
    cryptPassword.arFlag(fois) = 0
Next fois

champ1 = Key
lngchamp1 = Len(champ1)

For fois = 1 To lngchamp1
    cle(fois) = Asc(Mid(champ1, fois, 1))
    cryptPassword.arFlag(fois) = cle(fois)
Next fois

'For currentCounter = 1 To 16
'    cryptPassword.arFlag(currentCounter) = cryptPassword.arFlag(currentCounter) Xor tmpAsc
'Next

champ1 = inp
lngchamp1 = Len(champ1)
For fois = 1 To lngchamp1
c = Asc(Mid(champ1, fois, 1))

Assemble

cfc = (((inter / 256) * 256) - (inter Mod 256)) / 256
cfd = inter Mod 256

For compte = 1 To 16

cle(compte) = cle(compte) Xor c

Next compte

c = c Xor (cfc Xor cfd)

d = (((c / 16) * 16) - (c Mod 16)) / 16
e = c Mod 16

cryptPassword.encodedPassword = cryptPassword.encodedPassword + Chr$(&H61 + d) ' d+&h61 give one letter range from a to p for the 4 high bits of c
cryptPassword.encodedPassword = cryptPassword.encodedPassword + Chr$(&H61 + e) ' e+&h61 give one letter range from a to p for the 4 low bits of c

Next fois

End Function

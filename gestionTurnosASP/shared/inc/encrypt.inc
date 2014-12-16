<%
'Private Function Decrypt(ByVal strEncryptionKey, ByVal strTextToEncrypt)
'    Dim outer, inner, Key, strTemp
'    For outer = 1 To Len(strEncryptionKey)
'        key = Asc(Mid(strEncryptionKey, outer, 1))
'        For inner = 1 To Len(strTextToEncrypt)
'            strTemp = strTemp & Chr(Asc(Mid(strTextToEncrypt, inner, 1)) Xor key)
'            key = (key + Len(strEncryptionKey)) Mod 256
'        Next
'        strTextToEncrypt = strTemp
'        strTemp = ""
'    Next
'    Decrypt = strTextToEncrypt
'End Function


Private Function Decrypt(ByVal sClave, ByVal sOriginal, blnAccion)
Dim LenOri 
Dim LenClave
Dim i, j 
Dim cO, cC 
Dim k 
Dim v 


LenOri = Len(sOriginal)
LenClave = Len(sClave)

v = ""
i = 0
For j = 1 To LenOri
    i = i + 1
    If i > LenClave Then
        i = 1
    End If
    cO = Asc(Mid(sOriginal, j, 1))
    cC = Asc(Mid(sClave, i, 1))
    If blnAccion Then
        k = cO + cC
        If k > 255 Then
            k = k - 255
        End If
    Else
        k = cO - cC
        If k < 0 Then
            k = k + 255
        End If
    End If
    v = v & Chr(k)
Next

Decrypt = v
End Function

%>

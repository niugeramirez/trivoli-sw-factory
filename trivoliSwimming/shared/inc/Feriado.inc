<%  on error goto 0

Dim nroConv 
Dim NombreConvenio 
Dim nroSuc 
Dim NombreSucursal 
Dim nroEstr 
Dim NombreEstr
Dim StrSql

Function Feriado(dia, Ternro)

    Feriado = False

    'Determino el nro de la estructura
    StrSql = " SELECT * FROM his_estructura "
    StrSql = StrSql & " INNER JOIN estructura ON his_estructura.estrnro = estructura.estrnro "
    StrSql = StrSql & " INNER JOIN Alcance_Testr ON his_estructura.tenro = Alcance_Testr.tenro "
    StrSql = StrSql & " WHERE Alcance_Testr.tanro = " & lngAlcanGrupo & " And " & _
             " his_estructura.Ternro = " & Ternro & " AND htetdesde <= " & cambiafecha(dia,"YMD",true) & _
             " AND (htethasta >= " & cambiafecha(dia,"YMD",true) & " Or htethasta Is Null )" & _
             " ORDER BY htetdesde DESC "

    rsOpen l_rs, cn, StrSql, 0 

    If Not l_rs.EOF Then
        nroEstr = l_rs("estrnro")
        NombreEstr = Trim(" " & l_rs("estrdabr"))
    End If
    l_rs.close
	
    StrSql = "SELECT * FROM Feriado WHERE ferifecha = " & cambiafecha(dia,"YMD",true) 
    rsOpen l_rs2, cn, StrSql, 0 
    
    If Not l_rs2.EOF Then
        If CInt(l_rs2("tipferinro")) = 1 Then
            'Pais
            StrSql = "SELECT * FROM Pais WHERE paisdef = -1 "
            rsOpen l_rs, cn, StrSql, 0 

            If Not l_rs.EOF Then
                If CInt(l_rs2("fericodext")) = CInt(l_rs("paisnro")) Then 
				   Feriado = True
				end if
            End If
   	        l_rs.close
        Else
            StrSql = "SELECT * FROM Fer_estr WHERE estrnro = " & nroEstr & " AND ferinro = " & l_rs2("ferinro")
            rsOpen l_rs, cn, StrSql, 0 
            If Not l_rs.EOF Then 
			   Feriado = True
			end if
   	        l_rs.close
        End If
    End If
    
	l_rs2.close
end Function 'Feriado(dia, Ternro)

%>

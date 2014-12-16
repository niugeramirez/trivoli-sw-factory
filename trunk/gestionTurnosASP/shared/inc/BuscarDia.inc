<%

Dim blnTrabaja  
Dim Ordendia 
Dim Nro_dia 
Dim blnDia_libre 
Dim Nro_subturno 
Dim NombreSTurno 
Dim Grupo 

Public Sub Buscar_Dia(fecha , Fecha_inicio , nro_turno , Ternro, p_asignacion)
Dim num_dia 
Dim dif_dias 
dim StrSql

    blnTrabaja = False
    Ordendia = 0
    Nro_dia = 0
    blnDia_libre = False
    Nro_subturno = 0
    NombreSTurno = ""
	
    StrSql = "SELECT * FROM gti_turno WHERE turnro = " & nro_turno
    if l_rs.state <> 0 then l_rs.close
    l_rs.open strsql,cn
        
    blnTrabaja = False
    Ordendia = -1 
    Nro_dia = -1 
    dif_dias = DateDiff("d", Fecha_inicio, fecha) + 1
    num_dia = dif_dias Mod l_rs("turtamanio")
    If (num_dia = 0) Then num_dia = l_rs("turtamanio")  'es el primer dia del turno */
    

    StrSql = "SELECT * FROM gti_detturtemp WHERE (ternro = " & Ternro & ") AND " & _
             " (gttempdesde <= " & cambiafecha(fecha,"YMD",true) & ") AND " & _
             " (" & cambiafecha(fecha,"YMD",true) & " <= gttemphasta)"
    if l_rs.state <> 0 then l_rs.close
    l_rs.open strsql,cn
    If Not l_rs.EOF Then blnDia_libre = l_rs("ttemplibre")

    
    ' Buscar el dia Correspondiente
    
    StrSql = "SELECT gti_dias.*,gti_subturno.subturdesabr FROM gti_subturno INNER JOIN gti_dias ON (gti_subturno.subturnro = gti_dias.subturnro) WHERE " & _
             " (turnro = " & nro_turno & ") AND (gti_dias.diaorden <= " & num_dia & ") ORDER BY diaorden DESC "
    
    if l_rs.state <> 0 then l_rs.close
    	l_rs.open strsql,cn
    	
    If Not l_rs.EOF Then
        blnTrabaja = True
        Ordendia = l_rs("diaorden")
        Nro_dia = l_rs("dianro")
        Nro_subturno = l_rs("subturnro")
        NombreSTurno = l_rs("subturdesabr")
        blnDia_libre = l_rs("Dialibre")

	If p_asignacion Then
             StrSql = "SELECT * FROM gti_detturtemp WHERE (ternro = " & Ternro & ") AND " & _
                	 " (gttempdesde <= " & cambiafecha(fecha,"YMD",true) & ") AND " & _
	                 " (" & cambiafecha(fecha,"YMD",true) & " <= gttemphasta)"
             if l_rs.state <> 0 then l_rs.close
	     l_rs.open strsql,cn
       	     if Not l_rs.EOF Then blnDia_libre = l_rs("ttemplibre")
        End If
    End if	

End Sub


%>
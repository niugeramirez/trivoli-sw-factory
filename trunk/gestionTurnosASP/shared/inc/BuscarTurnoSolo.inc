<%
dim Tiene_Turno

Sub BuscarTurno(l_ternro, l_fecha)
on error goto 0
Dim l_StrSql
'Dim l_rs             ' Utilizo el objeto externo para no tener que crearlo cada vez
'                      Esto mejora el tiempo del proceso bastante 
'	Set l_rs = Server.CreateObject("ADODB.RecordSet")

    l_StrSql = " SELECT * FROM his_estructura "
    l_StrSql = l_StrSql & " INNER JOIN Alcance_Testr ON his_estructura.tenro = Alcance_Testr.tenro "
    l_StrSql = l_StrSql & " WHERE (tanro = " & lngAlcanGrupo & ") AND (ternro = " & l_ternro & ") AND " & _
             " (htetdesde <= " & cambiafecha(l_fecha,"YMD", true) & ") AND " & _
             " ((" & cambiafecha(l_fecha,"YMD", true) & " <= htethasta) or (htethasta is null))"

	rsOpen l_rs, cn, l_StrSql, 0 
	
	Tiene_Turno = not l_rs.eof
	
	l_rs.close
'	set l_rs = Nothing
'	response.write l_strsql

'	response.write ("<script>alert('" & l_strsql & "');</script>")
'	response.write ("<script>alert('" & Tiene_Turno & "');</script>")

end sub
%>



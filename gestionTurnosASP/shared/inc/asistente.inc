
<%
'Archivo: asistente.inc
'Descripción: Funciones auxiliares del asistente de abm complejos
'Autor : Alvaro Bayon
'Fecha: 4/11/2003
'Modificado:


 sub actualizarPasos(paso, externa,estado)
 ' Actualiza la tabla de pasos
 ' EN EL LLAMADOR DEBE INCLUIRSE adovbs.inc, fecha.inc
 Dim l_sql
 Dim rs
 Dim l_cm
	 set rs = Server.CreateObject("ADODB.RecordSet")
	 set l_cm = Server.CreateObject("ADODB.Command")
	 l_cm.activeconnection = Cn
	' Actualizo la tabla de pasos
	l_sql = "SELECT pasnro FROM paso_ext "
	l_sql = l_sql & " WHERE pasnro = " & paso & " AND extnro = " & externa
	rsOpenCursor rs, cn, l_sql,0, adOpenKeyset
	if rs.eof then
		l_sql = "INSERT INTO paso_ext "
		l_sql = l_sql & "(pasnro, extnro, extestado, extfecha, extusuario) "
		l_sql = l_sql & " VALUES (" & paso & "," & externa & "," & estado & "," & cambiafecha(date(),"YMD",true) & ",'" & session("username") & "')"
		l_cm.CommandText = l_sql
		cmExecute l_cm, l_sql, 0
	else
		l_sql = "UPDATE paso_ext"
		l_sql = l_sql & " SET extfecha = " & cambiafecha(date(),"YMD",true)
		l_sql = l_sql & ", extusuario = '" & session("username") & "', extestado = " & estado
		l_sql = l_sql & " WHERE pasnro = " & paso & " AND extnro = " & externa
		l_cm.CommandText = l_sql
		cmExecute l_cm, l_sql, 0
	end if
	rs.close
	set rs = nothing
	set l_cm = nothing
 end sub

%>
<% Option Explicit %>
<!--#include virtual="/serviciolocal/shared/inc/sec.inc"-->
<!--#include virtual="/serviciolocal/shared/inc/const.inc"-->
<!--#include virtual="/serviciolocal/shared/db/conn_db.inc"-->
<% 

Dim l_menuraiz
Dim l_menuorder
Dim l_nombre
Dim l_pagina
Dim l_accesos
Dim l_tipo
Dim l_cm
Dim l_sql


l_tipo = request.querystring("tipo")
l_menuraiz = request.form("menuraiz")
l_menuorder = request.form("menuorder")
l_nombre = ucase(request.form("nombre"))
l_pagina = ucase(request.form("pagina"))
l_accesos = request.form("menuaccess")

set l_cm = Server.CreateObject("ADODB.Command")
if l_tipo = "A" then 
	l_sql = "insert into menubtn "
	l_sql = l_sql & "(menuraiz, menuorder, btnpagina, btnnombre, btnaccess) "
	l_sql = l_sql & "values (" & l_menuraiz & ", " & l_menuorder & ", '" & l_pagina & "', '" & l_nombre & "','" & l_accesos & "')"
else
	l_sql = "update menubtn "
	l_sql = l_sql & "set btnaccess = '" & l_accesos & "' "
	l_sql = l_sql & " where menuraiz = " & l_menuraiz & " and menuorder = " & l_menuorder & " and btnpagina = '" & l_pagina & "' and btnnombre = '" & l_nombre & "'"
end if

l_cm.activeconnection = Cn
l_cm.CommandText = l_sql
Dim Buffer
on error resume next
l_cm.Execute
if err then
	Buffer = "<script>"
	if debug then 
		Buffer = Buffer & "alert(" & chr(34) & "Debug: TRUE\n"
		Buffer = Buffer & "Archivo: " & Request.ServerVariables("SCRIPT_NAME") & "\n"
		Buffer = Buffer & "Numero Error: " & err.number & "\n"
		Buffer = Buffer & "Descripcion: " & err.description & "\n"
		Buffer = Buffer & "SQL: " & l_sql & chr(34) & ");"
		Buffer = Buffer & "prompt('SQL String:',"&chr(34)&l_sql&chr(34)&");"
	else
		Buffer = Buffer & "alert('" & err.description & "');"
	end if
	Buffer = Buffer & "</script>"
	response.write Buffer
	response.end
end if

Set cn = Nothing
Set l_cm = Nothing

Response.write "<script>alert('Operación Realizada.');window.opener.ifrm.location = 'Armado_menu_11.asp?menuraiz=" & l_menuraiz & "&menuorder=" & l_menuorder & "';window.close();</script>"
%>

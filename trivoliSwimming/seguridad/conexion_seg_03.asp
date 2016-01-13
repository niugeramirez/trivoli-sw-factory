<% Option Explicit %>
<!--#include virtual="/turnos/shared/inc/sec.inc"-->
<!--#include virtual="/turnos/shared/inc/const.inc"-->
<!--#include virtual="/turnos/shared/db/conn_db.inc"-->
<% 
'Archivo: conexion_seg_00.asp
'Descripción: 
'Autor: Lisandro Moro
'Fecha: 15/03/2005
'Modificado:
on error goto 0

Dim l_tipo
Dim l_cm
Dim l_sql
Dim l_cnnro
Dim l_cndesc
Dim l_cnstring

l_tipo = request.querystring("tipo")

l_cnnro = request.Form("cnnro")
l_cndesc = request.Form("cndesc")
l_cnstring = request.Form("cnstring")

set l_cm = Server.CreateObject("ADODB.Command")
if l_tipo = "A" then 
	l_sql = "insert into conexion "
	l_sql = l_sql & "(cndesc, cnstring) "
	l_sql = l_sql & "values ('" & l_cndesc & "','" & l_cnstring & "')"
else
	l_sql = "update conexion "
	l_sql = l_sql & "set cndesc = '" & l_cndesc & "',"
	l_sql = l_sql & "cnstring = '" & l_cnstring & "'"
	l_sql = l_sql & " where cnnro = " & l_cnnro
end if

Dim Buffer
on error resume next
l_cm.activeconnection = Cn
l_cm.CommandText = l_sql
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
	Buffer = Buffer & "window.close();"
	Buffer = Buffer & "</script>"
	response.write Buffer
	response.end
end if
Set cn = Nothing
Set l_cm = Nothing
Response.write "<script>alert('Operación Realizada.');window.opener.ifrm.location = 'conexion_seg_01.asp?orden='+window.opener.ifrm.datos.orden.value;window.close();</script>"
%>

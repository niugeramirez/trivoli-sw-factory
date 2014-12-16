<%  Option Explicit %>
<!--#include virtual="/serviciolocal/shared/inc/sec.inc"-->
<!--#include virtual="/serviciolocal/shared/inc/const.inc"-->
<!--#include virtual="/serviciolocal/shared/db/conn_db.inc"-->
<% 
' Declaracion de Variables locales  -------------------------------------
Dim l_tipo
Dim l_cm
Dim l_sql
dim l_rs

Dim l_caudnro
Dim l_cauddes
dim	l_caudact

' traer valores del form de alta/modificacion -------------------------------------

l_tipo = request("tipo")

l_caudnro			= Request.form("caudnro")
l_cauddes			= Request.form("cauddes")
l_caudact			= Request.form("caudact")

' trasnformar valor de checkboxes en valores logicos --------------------------
IF l_caudact = "on" then
' informix -1 =======================
	l_caudact = -1
else
	l_caudact = 0
end if
	
set l_cm = Server.CreateObject("ADODB.Command")
if l_tipo = "A" then 
	Set l_rs = Server.CreateObject("ADODB.RecordSet")
	l_sql = "SELECT caudnro, "  
	l_sql = l_sql & " cauddes,  "
	l_sql = l_sql & " caudact   "
	l_sql = l_sql & " FROM confaud"
	l_sql = l_sql & " WHERE cauddes = '" & trim(l_cauddes) & "'"
	l_rs.MaxRecords = 1
	rsOpen l_rs, cn, l_sql, 0 
	if not l_rs.eof then
		l_rs.Close
		set l_rs = nothing
		Response.write "<script>alert('Existe otra Configuracion de Auditoria con esta descripcion.');window.close();</script>"
	else
		l_rs.Close
		set l_rs = nothing
		l_sql = "INSERT INTO confaud "
		l_sql = l_sql & "(cauddes , caudact) "
		l_sql = l_sql & " values ('" 
		l_sql = l_sql & l_cauddes & "'," 
		l_sql = l_sql & l_caudact & ")"
	end if	
else
	Set l_rs = Server.CreateObject("ADODB.RecordSet")
	l_sql = "SELECT caudnro, "  
	l_sql = l_sql & " cauddes,  "
	l_sql = l_sql & " caudact   "
	l_sql = l_sql & " FROM confaud"
	l_sql = l_sql & " WHERE cauddes = '" & trim(l_cauddes) & "'"
	l_sql = l_sql & " AND   caudnro <> " & l_caudnro
	l_rs.MaxRecords = 1
	rsOpen l_rs, cn, l_sql, 0 
	if not l_rs.eof then
		l_rs.Close
		set l_rs = nothing
		Response.write "<script>alert('Existe otra Configuracion de Auditoria con esta descripcion.');window.close();</script>"
	else
		l_rs.Close
		set l_rs = nothing
		l_sql = "UPDATE confaud SET "
		l_sql = l_sql & "cauddes	= '"  & l_cauddes  & "',"
		l_sql = l_sql & "caudact	= "   & l_caudact   
		l_sql = l_sql & " WHERE caudnro = " & l_caudnro
	end if
end if

l_cm.activeconnection = Cn
l_cm.CommandText = l_sql
cmExecute l_cm, l_sql, 0	

Response.write "<script>alert('Operación Realizada.');window.opener.ifrm.location = 'configuracion_auditoria_01.asp';window.close();</script>"
%>

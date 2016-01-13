<%  Option Explicit %>
<!--#include virtual="/turnos/shared/inc/sec.inc"-->
<!--#include virtual="/turnos/shared/inc/const.inc"-->
<!--#include virtual="/turnos/shared/db/conn_db.inc"-->
<% 
' Declaracion de Variables locales  -------------------------------------
Dim l_tipo
Dim l_cm
Dim l_sql
dim l_rs

Dim l_repnro
Dim l_repdesc
dim	l_repagr

' traer valores del form de alta/modificacion -------------------------------------

l_tipo = request("tipo")

l_repnro   = Request.form("repnro")
l_repdesc  = Request.form("repdesc")
l_repagr   = Request.form("repagr")

' trasnformar valor de checkboxes en valores logicos --------------------------
IF l_repagr = "on" then
' informix -1 =======================
	l_repagr = -1
else
	l_repagr = 0
end if
	
set l_cm = Server.CreateObject("ADODB.Command")
if l_tipo = "A" then 
	Set l_rs = Server.CreateObject("ADODB.RecordSet")
	l_sql = "SELECT repnro, "  
	l_sql = l_sql & " repdesc,  "
	l_sql = l_sql & " repagr   "
	l_sql = l_sql & " FROM reporte"
	l_sql = l_sql & " WHERE repdesc = '" & trim(l_repdesc) & "'"
	l_rs.MaxRecords = 1
	rsOpen l_rs, cn, l_sql, 0 
	if not l_rs.eof then
		l_rs.Close
		set l_rs = nothing
		Response.write "<script>alert('Existe otro Reporte con esta descripcion.');window.close();</script>"
	else
		l_rs.Close
		set l_rs = nothing
		l_sql = "INSERT INTO reporte "
		l_sql = l_sql & "(repdesc , repagr) "
		l_sql = l_sql & " values ('" 
		l_sql = l_sql & l_repdesc & "'," 
		l_sql = l_sql & l_repagr & ")"
	end if	
else
	Set l_rs = Server.CreateObject("ADODB.RecordSet")
	l_sql = "SELECT repnro, "  
	l_sql = l_sql & " repdesc,  "
	l_sql = l_sql & " repagr   "
	l_sql = l_sql & " FROM reporte"
	l_sql = l_sql & " WHERE repdesc = '" & trim(l_repdesc) & "'"
	l_sql = l_sql & " AND   repnro <> " & l_repnro
	l_rs.MaxRecords = 1
	rsOpen l_rs, cn, l_sql, 0 
	if not l_rs.eof then
		l_rs.Close
		set l_rs = nothing
		Response.write "<script>alert('Existe otro Reporte con esta descripcion.');window.close();</script>"
	else
		l_rs.Close
		set l_rs = nothing
		l_sql = "UPDATE reporte SET "
		l_sql = l_sql & "repdesc	= '"  & l_repdesc  & "',"
		l_sql = l_sql & "repagr	= "   & l_repagr   
		l_sql = l_sql & " WHERE repnro = " & l_repnro
	end if
end if

l_cm.activeconnection = Cn
l_cm.CommandText = l_sql
cmExecute l_cm, l_sql, 0	

Response.write "<script>alert('Operación Realizada.');window.opener.ifrm.location = 'reporte_01.asp';window.close();</script>"
%>

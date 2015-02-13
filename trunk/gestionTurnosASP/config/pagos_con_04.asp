<% Option Explicit %>
<!--#include virtual="/turnos/shared/db/conn_db.inc"-->
<!--#include virtual="/turnos/shared/inc/const.inc"-->
<% 
'Archivo: companies_con_04.asp
'Descripción: ABM de Companies
'Autor : Raul Chinestra
'Fecha: 26/11/2007

'on error goto 0
Dim l_cm
Dim l_rs
Dim l_sql
Dim l_id
	
l_id = request.querystring("cabnro")
Set l_rs = Server.CreateObject("ADODB.RecordSet")
set l_cm = Server.CreateObject("ADODB.Command")
'l_sql = "SELECT counro"
'l_sql = l_sql & " FROM for_port "
'l_sql  = l_sql  & " WHERE counro = " & l_counro
'rsOpen l_rs, cn, l_sql, 0 
'if not l_rs.eof then
'	Response.write "<script>alert('Existen Ports asociados a este Country.\nNo es posible dar de baja.');window.close();</script>"
'else
	l_sql = " DELETE FROM pagos WHERE id = " & l_id
'end if
'l_rs.close

l_cm.activeconnection = Cn
l_cm.CommandText = l_sql
cmExecute l_cm, l_sql, 0

cn.Close
Set cn = Nothing
%>
<script>
	alert('Operación Realizada.');
	window.opener.ifrm.location.reload();
	window.close();
</script>
	





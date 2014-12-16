<% Option Explicit %>
<!--#include virtual="/ticket/shared/db/conn_db.inc"-->
<!--#include virtual="/ticket/shared/inc/const.inc"-->
<% 
'Archivo: terminal_con_04.asp
'Descripción: ABM de Terminales
'Autor : Gustavo Manfrin
'Fecha: 19/04/2005

'on error goto 0
Dim l_cm
Dim l_rs
Dim l_sql
Dim l_ternro
	
l_ternro = request.querystring("cabnro")
Set l_rs = Server.CreateObject("ADODB.RecordSet")
set l_cm = Server.CreateObject("ADODB.Command")
l_rs.close
'l_sql = "SELECT depnro"
'l_sql = l_sql & " FROM tkt_pro_dep"
'l_sql  = l_sql  & " WHERE depnro = " & l_depnro
'rsOpen l_rs, cn, l_sql, 0 
'if not l_rs.eof then
'	Response.write "<script>alert('Existen Productos asociados a este Depósito.\nNo es posible dar de baja.');window.close();</script>"
'else
	l_sql = " DELETE FROM tkt_terminal WHERE ternro = " & l_ternro
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
	





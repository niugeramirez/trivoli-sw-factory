<% Option Explicit %>
<!--#include virtual="/serviciolocal/shared/db/conn_db.inc"-->
<!--#include virtual="/serviciolocal/shared/inc/const.inc"-->
<% 
'Archivo: berths_con_04.asp
'Descripción: ABM de Berths
'Autor : Raul Chinestra
'Fecha: 23/11/2007

on error goto 0
Dim l_cm
Dim l_rs
Dim l_sql
Dim l_connro
	
l_connro = request.querystring("cabnro")
Set l_rs = Server.CreateObject("ADODB.RecordSet")
set l_cm = Server.CreateObject("ADODB.Command")
'l_sql = "SELECT bernro"
'l_sql = l_sql & " FROM for_berth "
'l_sql  = l_sql  & " WHERE bernro = " & l_bernro
'rsOpen l_rs, cn, l_sql, 0 
'if not l_rs.eof then
'	Response.write "<script>alert('Existen Countries asociados a este Area.\nNo es posible dar de baja.');window.close();</script>"
'else
	l_sql = " DELETE FROM buq_contenido WHERE connro = " & l_connro
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
	





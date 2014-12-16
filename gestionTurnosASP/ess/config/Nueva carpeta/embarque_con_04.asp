<% Option Explicit %>
<!--#include virtual="/serviciolocal/shared/db/conn_db.inc"-->
<!--#include virtual="/serviciolocal/shared/inc/const.inc"-->
<% 
'Archivo: embarque_con_04.asp
'Descripción: Abm de embarque
'Autor : Gustavo Manfrin
'Fecha: 18/09/2006

'on error goto 0
Dim l_cm
Dim l_rs
Dim l_sql
Dim l_embnro
	
l_embnro = request.querystring("cabnro")
Set l_rs = Server.CreateObject("ADODB.RecordSet")
set l_cm = Server.CreateObject("ADODB.Command")
	l_cm.activeconnection = Cn
l_sql = "SELECT embnro"
l_sql = l_sql & " FROM tkt_asiemb"
l_sql  = l_sql  & " WHERE embnro = " & l_embnro
rsOpen l_rs, cn, l_sql, 0 
if not l_rs.eof then
	Response.write "<script>alert('Existen tarjetas asociados a este Embarque.\nNo es posible dar de baja.');window.close();</script>"
else
	l_rs.close
	l_sql = " DELETE FROM tkt_embarque WHERE embnro = " & l_embnro
	l_cm.activeconnection = Cn
	l_cm.CommandText = l_sql
	cmExecute l_cm, l_sql, 0
end if
cn.Close
Set cn = Nothing
%>
<script>
	alert('Operación Realizada.');
	window.opener.ifrm.location.reload();
	window.close();
</script>
	





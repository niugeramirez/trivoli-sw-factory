<% Option Explicit %>
<!--#include virtual="/serviciolocal/shared/db/conn_db.inc"-->
<!--#include virtual="/serviciolocal/shared/inc/const.inc"-->
<% 
'Archivo: impresoras_con_04.asp
'Descripción: ABM de Impresoras
'Autor : Lisandro Moro
'Fecha: 26/09/2005
'Modificado: 

'on error goto 0
Dim l_cm
Dim l_rs
Dim l_sql
Dim l_impnro
	
l_impnro = request.querystring("cabnro")
Set l_rs = Server.CreateObject("ADODB.RecordSet")
set l_cm = Server.CreateObject("ADODB.Command")
l_rs.close
l_sql = " SELECT tkt_impresora.impnom "
l_sql = l_sql & " FROM tkt_terminal, tkt_impresora "
'l_sql = l_sql & " INNER JOIN tkt_impresora ON tkt_terminal.impnom = tkt_impresora.impnom  "
l_sql  = l_sql  & " WHERE impnro = " & l_impnro
l_sql  = l_sql  & " AND (impnom = terimptick OR impnom = terimpcpor OR impnom = terimpremi OR impnom = terimpetiq) "

rsOpen l_rs, cn, l_sql, 0 
if not l_rs.eof then
	Response.write "<script>alert('Existen Terminales asociados a esta Impresora.\nNo es posible dar de baja.');window.close();</script>"
else
	l_sql = " DELETE FROM tkt_impresora WHERE impnro = " & l_impnro
	l_cm.activeconnection = Cn
	l_cm.CommandText = l_sql
	cmExecute l_cm, l_sql, 0
end if
l_rs.close

cn.Close
Set cn = Nothing
%>
<script>
	alert('Operación Realizada.');
	window.opener.ifrm.location.reload();
	window.close();
</script>
	





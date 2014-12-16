<% Option Explicit %>
<!--#include virtual="/serviciolocal/shared/db/conn_db.inc"-->
<!--#include virtual="/serviciolocal/shared/inc/const.inc"-->
<% 
'Archivo: entrec_con_04.asp
'Descripción: ABM Entregadores/recibidores
'Autor : Alvaro Bayon
'Fecha: 11/02/2005

'on error goto 0
Dim l_cm
Dim l_rs
Dim l_sql
Dim l_entnro
	
'l_entnro = request.querystring("cabnro")
'set l_cm = Server.CreateObject("ADODB.Command")
'l_sql = "UPDATE tkt_entrec"
'l_sql  = l_sql  & " SET entact = 0"
'l_sql  = l_sql  & " WHERE entnro = " & l_entnro


l_entnro = request.querystring("cabnro")
Set l_rs = Server.CreateObject("ADODB.RecordSet")
set l_cm = Server.CreateObject("ADODB.Command")
'Valido que el entregador no este en ninguna carta de porte
l_sql = "SELECT entnro"
l_sql = l_sql & " FROM tkt_cartaporte"
l_sql  = l_sql  & " WHERE entnro = " & l_entnro
rsOpen l_rs, cn, l_sql, 0 
if not l_rs.eof then
	Response.write "<script>alert('Existen Cartas de Porte asociadas a este Entregador/Recibidor.\nNo es posible dar de baja.');window.close();</script>"
else
	l_rs.close
	'Valido que el entregador(tipternro=3) no tenga cargados documentos
	l_sql = "SELECT valnro"
	l_sql = l_sql & " FROM tkt_terdoc"
	l_sql  = l_sql  & " WHERE tipternro = 3"
	l_sql  = l_sql  & " AND valnro = " & l_entnro
	rsOpen l_rs, cn, l_sql, 0 
	if not l_rs.eof then
		Response.write "<script>alert('Existen Documentos asociados a este Entregador/Recibidor.\nNo es posible dar de baja.');window.close();</script>"
	else
		l_sql = " DELETE FROM tkt_entrec WHERE entnro = " & l_entnro
	end if
	l_rs.close
end if
l_cm.activeconnection = Cn
l_cm.CommandText = l_sql
cmExecute l_cm, l_sql, 0

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
	





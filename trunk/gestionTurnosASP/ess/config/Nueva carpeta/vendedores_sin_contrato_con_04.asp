<% Option Explicit %>
<!--#include virtual="/ticket/shared/inc/sec.inc"-->
<!--#include virtual="/ticket/shared/inc/const.inc"-->
<!--#include virtual="/ticket/shared/db/conn_db.inc"-->
<%
'Archivo: vendedores_sin_contrato_con_04.asp
'Descripción: Eliminar Vendedores habilitados a descargar sin contrato
'Autor : Lisandro Moro
'Fecha: 11/02/2005

Dim l_rs
Dim l_cm
Dim l_sql

Dim l_vencornro
Dim l_pronro
Dim l_tipo
Dim l_planro

on error goto 0 
l_vencornro = request.QueryString("cabnro")
l_pronro	= request.QueryString("pronro")
l_tipo	= request.QueryString("tipo")

set l_cm = Server.CreateObject("ADODB.Command")
Set l_rs = Server.CreateObject("ADODB.RecordSet")
l_sql = "SELECT tkt_config.lugnro, planro "
l_sql = l_sql & " FROM tkt_config "
l_sql = l_sql & " INNER JOIN tkt_lugar ON tkt_lugar.lugnro = tkt_config.lugnro "
rsOpen l_rs, cn, l_sql, 0 
if not l_rs.eof then
	l_planro = l_rs("planro")
else
	'Error, falta configurar la planta o el lugar
	Response.Write "error"
end if
l_rs.close
'Realizo la validacio que no exista otro con el mismo vendedor, producto y planta
l_sql = "SELECT vencornro, planro, pronro "
l_sql = l_sql & " FROM tkt_autsincon "
l_sql = l_sql & " WHERE vencornro = " & l_vencornro & " AND planro = " & l_planro & " AND pronro = " & l_pronro
rsOpen l_rs, cn, l_sql, 0 
if not l_rs.eof then
		l_sql = " DELETE FROM tkt_autsincon "
		l_sql = l_sql & " WHERE vencornro = " & l_vencornro & " AND planro = " & l_planro & " AND pronro = " & l_pronro
		l_cm.activeconnection = Cn
		l_cm.CommandText = l_sql
		cmExecute l_cm, l_sql, 0
else
	Response.Write "<script>alert('No existe el vendedor autorizado a descargar el producto.');</script>"
	Response.Write "<script>window.close();</script>"
	Response.End
end if
Set l_cm = Nothing

Response.write "<script>alert('Operación Realizada.');window.opener.ifrm.location.reload();window.close();</script>"
%>

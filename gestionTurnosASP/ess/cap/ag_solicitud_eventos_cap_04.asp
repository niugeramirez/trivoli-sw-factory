<% Option Explicit %>
<!--#include virtual="/serviciolocal/shared/db/conn_db.inc"-->
<!--#include virtual="/serviciolocal/shared/inc/const.inc"-->
<!--
Archivo: ag_solicitud_eventos_cap_04.asp
Descripción: Abm de Solicitud de Eventos
Autor : Raul Chinestra
Fecha: 30/04/2004
-->
<% 
Dim l_cm
Dim l_rs
Dim l_sql
Dim l_solnro
	
	l_solnro = request.querystring("cabnro")

	set l_cm = Server.CreateObject("ADODB.Command")
	l_sql = "DELETE FROM cap_solicitud WHERE solnro = " & l_solnro
	l_cm.activeconnection = Cn
	l_cm.CommandText = l_sql
	cmExecute l_cm, l_sql, 0
	
	'l_cm.close
	cn.Close
	Set cn = Nothing
	Response.write "<script>alert('Operación Realizada.');window.opener.ifrm.location.reload();window.close();</script>"
 	
%>

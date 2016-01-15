<% Option Explicit %>
<!--#include virtual="/trivoliSwimming/shared/db/conn_db.inc"-->
<!--#include virtual="/trivoliSwimming/shared/inc/const.inc"-->
<% 
Dim l_cm
Dim l_sql
Dim l_rs

Dim l_repnro

l_repnro			= Request.QueryString("repnro")

' borrar todas las columnas del reporte
set l_cm = Server.CreateObject("ADODB.Command")
l_sql = "DELETE FROM confrep WHERE repnro = " & l_repnro
l_cm.activeconnection = Cn
l_cm.CommandText = l_sql
cmExecute l_cm, l_sql, 0	

set l_cm = Server.CreateObject("ADODB.Command")
l_sql = "DELETE FROM reporte WHERE repnro = " & l_repnro
l_cm.activeconnection = Cn
l_cm.CommandText = l_sql
cmExecute l_cm, l_sql, 0	

Response.write "<script>alert('Operación Realizada.');window.opener.ifrm.location = 'reporte_01.asp';window.close();</script>"
%>

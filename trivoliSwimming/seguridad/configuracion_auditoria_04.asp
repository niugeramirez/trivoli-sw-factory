<% Option Explicit %>
<!--#include virtual="/turnos/shared/db/conn_db.inc"-->
<!--#include virtual="/turnos/shared/inc/const.inc"-->
<% 
Dim l_cm
Dim l_sql
Dim l_rs

Dim l_caudnro
Dim l_cauddes
dim	l_caudact

l_caudnro			= Request.QueryString("caudnro")

set l_cm = Server.CreateObject("ADODB.Command")
l_sql = "DELETE FROM confaud WHERE caudnro = " & l_caudnro
l_cm.activeconnection = Cn
l_cm.CommandText = l_sql
cmExecute l_cm, l_sql, 0	

Response.write "<script>alert('Operación Realizada.');window.opener.ifrm.location = 'configuracion_auditoria_01.asp';window.close();</script>"
%>

<% Option Explicit %>
<!--#include virtual="/trivoliSwimming/shared/db/conn_db.inc"-->
<!--#include virtual="/trivoliSwimming/shared/inc/const.inc"-->
<%
'Archivo        : conf_mil_04.asp
'Descripcion    : Modulo que se encarga de admin. los servidores de mail
'Creador        : Lisandro Moro
'Fecha Creacion : 08/03/2005
'Modificacion   :

Dim l_cm
Dim l_sql

Dim l_cfgemailnro

l_cfgemailnro = request.QueryString("cfgemailnro")

set l_cm = Server.CreateObject("ADODB.Command")
l_sql = "DELETE FROM conf_email WHERE cfgemailnro = " & l_cfgemailnro
l_cm.activeconnection = Cn
l_cm.CommandText = l_sql
cmExecute l_cm, l_sql, 0	

Response.write "<script>alert('Operación Realizada.');window.opener.ifrm.location.reload();window.close();</script>"

%>

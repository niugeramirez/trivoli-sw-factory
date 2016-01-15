<% Option Explicit %>
<!--#include virtual="/trivoliSwimming/shared/db/conn_db.inc"-->
<!--#include virtual="/trivoliSwimming/shared/inc/const.inc"-->
<!--
-----------------------------------------------------------------------------
Archivo        : conf_x_empr_04.asp
Descripcion    : Modulo que se encarga de las Bajas de conf de empresas.
Creador        : Scarpa D.
Fecha Creacion : 21/08/2003
-----------------------------------------------------------------------------
-->
<% 
Dim l_cm
Dim l_sql

Dim l_confnro

l_confnro = Request.QueryString("confnro")

set l_cm = Server.CreateObject("ADODB.Command")
l_sql = "DELETE FROM confper WHERE confnro = " & l_confnro
l_cm.activeconnection = Cn
l_cm.CommandText = l_sql
cmExecute l_cm, l_sql, 0	

Response.write "<script>alert('Operación Realizada.');window.opener.ifrm.location.reload();window.close();</script>"
%>

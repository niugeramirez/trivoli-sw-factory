<% Option Explicit %>
<!--#include virtual="/serviciolocal/shared/db/conn_db.inc"-->
<!--#include virtual="/serviciolocal/shared/inc/const.inc"-->
<!--
Archivo    : vales_liq_04.asp
Descripción: ABM de vales - Baja
Autor      : Scarpa D. 
Fecha      : 12/01/2004
  Modificado  : 12/09/2006 Raul Chinestra - se agregó Vales en Autogestión   
-->
<% 
on error goto 0

Dim l_cm
Dim l_rs
Dim l_sql
Dim l_valnro
	
l_valnro = request.querystring("valnro")

set l_cm = Server.CreateObject("ADODB.Command")
l_sql = "DELETE FROM vales WHERE valnro = " & l_valnro

l_cm.activeconnection = Cn
l_cm.CommandText = l_sql
cmExecute l_cm, l_sql, 0
			
cn.Close
Set cn = Nothing

Response.write "<script>alert('Operación Realizada.');window.opener.ifrm.location.reload();window.close();</script>"

%>

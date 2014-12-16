<% Option Explicit %>
<!--#include virtual="/serviciolocal/shared/db/conn_db.inc"-->

<% 
'---------------------------------------------------------------------------------
'Archivo	: nota_adp_04.asp
'Descripción: Baja de notas
'Autor		: Claudia Cecilia Rossi
'Fecha		: 30-08-2003
'Modificado	: parametro notanro (nuevo campo)
'----------------------------------------------------------------------------------
%>
<script>
var xc = screen.availWidth;
var yc = screen.availHeight;
window.moveTo(xc,yc);	
window.resizeTo(20,20);</script>
<%
Dim l_cm
Dim l_sql
Dim l_rs
Dim l_ternro
Dim l_notanro

l_notanro   = request.QueryString("notanro")
'l_ternro    = request.QueryString("ternro")

	set l_cm = Server.CreateObject("ADODB.Command")
	l_sql = "DELETE FROM notas_ter " 
	l_sql = l_sql & " WHERE notanro  =" & l_notanro
	l_cm.activeconnection = Cn
	l_cm.CommandText = l_sql
	cmExecute l_cm, l_sql, 0
	cn.Close
	Set cn = Nothing
	Response.write "<script>alert('Operación Realizada.');window.opener.ifrm.location.reload();window.close();</script>"

%>

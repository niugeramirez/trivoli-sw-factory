<% Option Explicit %>
<!--#include virtual="/ticket/shared/db/conn_db.inc"-->
<% 
'Archivo: sitimpositivas_con_02.asp
'Descripción: Abm de Situaciones Impositivas
'Autor : Raul Chinestra
'Fecha: 08/02/2005

'Datos del formulario
Dim l_sitimpnro
Dim l_sitimpdes

'ADO
Dim l_sql
Dim l_rs

%>
<html>
<head>
<link href="/ticket/shared/css/tables3.css" rel="StyleSheet" type="text/css">
<meta http-equiv="Content-Type" content="text/html; charset=iso-8859-1">
<title><%= Session("Titulo")%>Situaciones Impositivas - Ticket</title>
</head>
<script src="/ticket/shared/js/fn_ayuda.js"></script>
<script src="/ticket/shared/js/fn_windows.js"></script>
<script>

</script>
<% 
Set l_rs = Server.CreateObject("ADODB.RecordSet")
l_sitimpnro = request.querystring("cabnro")
l_sql = "SELECT  sitimpdes "
l_sql = l_sql & " FROM tkt_sitimp "
l_sql  = l_sql  & " WHERE sitimpnro = " & l_sitimpnro
rsOpen l_rs, cn, l_sql, 0 
if not l_rs.eof then
	l_sitimpdes = l_rs("sitimpdes")
end if
l_rs.Close
%>
<body leftmargin="0" rightmargin="0" topmargin="0" bottommargin="0">
<form name="datos">
<table cellspacing="0" cellpadding="0" border="0" width="100%" height="100%">
<tr>
    <td class="th2"  nowrap>Situaciones Impositivas</td>
	<td class="th2" align="right">
		<a class=sidebtnHLP href="Javascript:ayuda('<%= Request.ServerVariables("SCRIPT_NAME")%>');">Ayuda</a>
	</td>
</tr>
<tr>
    <td height="100%" align="right"><b>Descripción:</b></td>
	<td height="100%">
		<input type="text" readonly class="deshabinp" name="sitimpdes" size="60" maxlength="50" value="<%= l_sitimpdes %>">
	</td>
</tr>
<tr>
    <td colspan="2" align="right" class="th2">
		<a class=sidebtnABM href="Javascript:window.close()">Salir</a>
	</td>
</tr>
</table>
</form>
<%
set l_rs = nothing
Cn.Close
set Cn = nothing
%>
</body>
</html>

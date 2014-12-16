<% Option Explicit %>
<!--#include virtual="/ticket/shared/db/conn_db.inc"-->
<% 
'Archivo: plantas_con_02.asp
'Descripción: Abm de Plantas de Trabajo
'Autor : Alvaro Bayon
'Fecha: 08/02/2005
'Modificado: 

'Datos del formulario
Dim l_planro
Dim l_plades
Dim l_locdes

'ADO
Dim l_sql
Dim l_rs

%>
<html>
<head>
<link href="/ticket/shared/css/tables3.css" rel="StyleSheet" type="text/css">
<meta http-equiv="Content-Type" content="text/html; charset=iso-8859-1">
<title><%= Session("Titulo")%>Plantas - Ticket</title>
</head>
<script src="/ticket/shared/js/fn_ayuda.js"></script>
<script src="/ticket/shared/js/fn_windows.js"></script>
<script>

</script>
<% 
Set l_rs = Server.CreateObject("ADODB.RecordSet")
l_planro = request.querystring("cabnro")
l_sql = "SELECT planro, plades, locdes"
l_sql = l_sql & " FROM tkt_planta "
l_sql = l_sql & " INNER JOIN tkt_localidad ON  tkt_localidad.locnro = tkt_planta.locnro"
l_sql  = l_sql  & " WHERE planro = " & l_planro
rsOpen l_rs, cn, l_sql, 0 
if not l_rs.eof then
	l_plades = l_rs("plades")
	l_locdes = l_rs("locdes")
end if
l_rs.Close
%>
<body leftmargin="0" rightmargin="0" topmargin="0" bottommargin="0">
<form name="datos">

<table cellspacing="0" cellpadding="0" border="0" width="100%" height="100%">
<tr>
    <td class="th2" nowrap>Plantas</td>
	<td class="th2" align="right">
		<a class=sidebtnHLP href="Javascript:ayuda('<%= Request.ServerVariables("SCRIPT_NAME")%>');">Ayuda</a>
	</td>
</tr>
<tr>
	<td colspan="2" height="100%">
		<table border="0" cellspacing="0" cellpadding="0">
			<tr>
				<td width="50%"></td>
				<td>
					<table cellspacing="0" cellpadding="0" border="0">
					<tr>
						<td align="right"><b>Descripción:</b></td>
						<td >
							<input type="text" readonly class="deshabinp" name="plades" size="60" maxlength="50" value="<%= l_plades %>">
						</td>
					</tr>
					<tr>
					    <td align="right"><b>Localidad:</b></td>
						<td >
							<input type="text" readonly class="deshabinp" name="locdes" size="60" maxlength="50" value="<%= l_locdes %>">
						</td>
					</tr>
					</table>
				</td>
				<td width="50%"></td>
			</tr>
		</table>
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
cn.Close
set cn = nothing
%>
</body>
</html>

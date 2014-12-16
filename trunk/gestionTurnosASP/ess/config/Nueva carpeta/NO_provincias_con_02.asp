<% Option Explicit %>
<!--#include virtual="/ticket/shared/db/conn_db.inc"-->
<% 
'Archivo: provincias_con_02.asp
'Descripción: Consulta de Provincias
'Autor : Alvaro Bayon
'Fecha: 08/02/2005
'Modificado: 

'Datos del formulario
Dim l_pronro
Dim l_procod
Dim l_prodes
Dim l_proobl

'ADO
Dim l_sql
Dim l_rs

%>
<html>
<head>
<link href="/ticket/shared/css/tables3.css" rel="StyleSheet" type="text/css">
<meta http-equiv="Content-Type" content="text/html; charset=iso-8859-1">
<title><%= Session("Titulo")%>Provincias - Ticket</title>
</head>
<script src="/ticket/shared/js/fn_ayuda.js"></script>
<script src="/ticket/shared/js/fn_windows.js"></script>
<script>

</script>
<% 
Set l_rs = Server.CreateObject("ADODB.RecordSet")
l_pronro = request.querystring("cabnro")
l_sql = "SELECT  procod, prodes, proobl "
l_sql = l_sql & " FROM tkt_provincia "
l_sql  = l_sql  & " WHERE pronro = " & l_pronro
rsOpen l_rs, cn, l_sql, 0 
if not l_rs.eof then
	l_procod = l_rs("procod")
	l_prodes = l_rs("prodes")
	l_proobl = l_rs("proobl")
end if
l_rs.Close
%>
<body leftmargin="0" rightmargin="0" topmargin="0" bottommargin="0">
<form name="datos">

<table cellspacing="0" cellpadding="0" border="0" width="100%" height="100%">
<tr>
    <td class="th2"  nowrap>Provincias</td>
	<td class="th2" align="right">
		<a class=sidebtnHLP href="Javascript:ayuda('<%= Request.ServerVariables("SCRIPT_NAME")%>');">Ayuda</a>
	</td>
</tr>
<tr>
	<td colspan="2" height="100%">
		<table cellpadding="0" cellspacing="0">
			<tr>
				<td width="50%"></td>
				<td>
					<table cellpadding="0" cellspacing="0">
						<tr>
						    <td align="right"><b>Código:</b></td>
							<td>
								<input type="text" readonly class="deshabinp" name="procod" size="10" maxlength="10" value="<%= l_procod %>">
							</td>
						</tr>
						<tr>
						    <td align="right"><b>Descripción:</b></td>
							<td>
								<input type="text" readonly class="deshabinp" name="prodes" size="60" maxlength="50" value="<%= l_prodes %>">
							</td>
						</tr>
						<tr>
						    <td align="right"><b>Oblea:</b></td>
							<td>
								<input type="Checkbox" name="proobl" disabled  <% if l_proobl = -1 then %>checked<% End If %>> 
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

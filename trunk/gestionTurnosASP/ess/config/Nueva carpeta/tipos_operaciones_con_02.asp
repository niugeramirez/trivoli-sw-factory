<% Option Explicit %>
<!--#include virtual="/ticket/shared/db/conn_db.inc"-->
<% 
'Archivo: tipos_operaciones_con_02.asp
'Descripción: Abm de Tipos de Operaciones
'Autor : Alvaro Bayon
'Fecha: 08/02/2005
'Modificado: 

'Datos del formulario
Dim l_tipopenro
Dim l_tipopedes

'ADO
Dim l_sql
Dim l_rs

%>
<html>
<head>
<link href="/ticket/shared/css/tables3.css" rel="StyleSheet" type="text/css">
<meta http-equiv="Content-Type" content="text/html; charset=iso-8859-1">
<title><%= Session("Titulo")%>Tipos de Operaciones - Ticket</title>
</head>
<script src="/ticket/shared/js/fn_ayuda.js"></script>
<script src="/ticket/shared/js/fn_windows.js"></script>
<script>

</script>
<% 
Set l_rs = Server.CreateObject("ADODB.RecordSet")
l_tipopenro = request.querystring("cabnro")
l_sql = "SELECT  tipopedes "
l_sql = l_sql & " FROM tkt_tipooperacion "
l_sql  = l_sql  & " WHERE tipopenro = " & l_tipopenro
rsOpen l_rs, cn, l_sql, 0 
if not l_rs.eof then
	l_tipopedes = l_rs("tipopedes")
end if
l_rs.Close
%>
<body leftmargin="0" rightmargin="0" topmargin="0" bottommargin="0">
<form name="datos">

<table cellspacing="0" cellpadding="0" border="0" width="100%" height="100%">
<tr>
    <td class="th2" nowrap>Tipos de Operaciones</td>
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
						    <td height="100%" align="right"><b>Descripción:</b></td>
							<td height="100%">
								<input type="text" readonly class="deshabinp" name="tipopedes" size="50" maxlength="50" value="<%= l_tipopedes %>">
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

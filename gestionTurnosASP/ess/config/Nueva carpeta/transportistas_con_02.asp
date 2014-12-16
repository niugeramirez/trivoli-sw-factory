<% Option Explicit %>
<!--#include virtual="/ticket/shared/db/conn_db.inc"-->
<% 
'Archivo: transportistas_con_02.asp
'Descripción: Abm de Transportistas
'Autor : Gustavo Manfrin
'Fecha: 09/02/2005
'Modificado: 

'Datos del formulario
Dim l_tranro
Dim l_tracod
Dim l_trades
Dim l_tradir
Dim l_tracaj
Dim l_tranrocaj
Dim l_prodes
Dim l_sitimpdes
Dim l_trarazsoc

'ADO
Dim l_sql
Dim l_rs

%>
<html>
<head>
<link href="/ticket/shared/css/tables3.css" rel="StyleSheet" type="text/css">
<meta http-equiv="Content-Type" content="text/html; charset=iso-8859-1">
<title><%= Session("Titulo")%>Transportistas - Ticket</title>
</head>
<script src="/ticket/shared/js/fn_ayuda.js"></script>
<script src="/ticket/shared/js/fn_windows.js"></script>
<script>

</script>
<% 
Set l_rs = Server.CreateObject("ADODB.RecordSet")
l_tranro = request.querystring("cabnro")

l_sql = "SELECT tranro, tracod, trades, trarazsoc, tracaj, tranrocaj, tradir, sitimpdes, prodes "
l_sql = l_sql & " FROM tkt_transportista "
l_sql = l_sql & " INNER JOIN tkt_provincia ON tkt_provincia.pronro = tkt_transportista.pronro  "
l_sql = l_sql & " INNER JOIN tkt_sitimp ON tkt_sitimp.sitimpnro = tkt_transportista.sitimpnro  "
l_sql = l_sql & " WHERE tranro = " & l_tranro

rsOpen l_rs, cn, l_sql, 0 
if not l_rs.eof then
	l_tranro = l_rs("tranro")
	l_tracod = l_rs("tracod")
	l_trades = l_rs("trades")
	l_tradir = l_rs("tradir")
	l_trarazsoc = l_rs("trarazsoc")
    l_tracaj = l_rs("tracaj")
	l_tranrocaj = l_rs("tranrocaj")	
	l_sitimpdes = l_rs("sitimpdes")
	l_prodes = l_rs("prodes")
end if
l_rs.Close
%>
<body leftmargin="0" rightmargin="0" topmargin="0" bottommargin="0">
<form name="datos">

<table cellspacing="0" cellpadding="0" border="0" width="100%" height="100%">
<tr>
    <td class="th2"  nowrap>Transportistas</td>
	<td class="th2" align="right">
		<a class=sidebtnHLP href="Javascript:ayuda('<%= Request.ServerVariables("SCRIPT_NAME")%>');">Ayuda</a>
	</td>
</tr>
<tr>
	<td colspan="2" width="100%" height="100%">	
		<table cellpadding="0" cellspacing="0">
			<tr>
				<td width="5%"></td>
				<td>
					<table cellpadding="0" cellspacing="0">
					<tr>
					    <td align="right"><b>Código:</b></td>
						<td>
							<input type="text" readonly class="deshabinp" name="tracod" size="20" maxlength="20" value="<%= l_tracod %>">
						</td>
					</tr>
					<tr>
					    <td align="right"><b>Nombre clave:</b></td>
						<td>
							<input type="text" readonly class="deshabinp" name="trades" size="60" maxlength="50" value="<%= l_trades %>">
						</td>
					</tr>
					<tr>
					    <td align="right"><b>Razón social:</b></td>
						<td>
							<input type="text" readonly class="deshabinp" name="trarazsoc" size="60" maxlength="50" value="<%= l_trarazsoc %>">
						</td>
					</tr>
					<tr>
					    <td align="right"><b>Dirección</b>:</b></td>
						<td>
							<input type="text" readonly class="deshabinp" name="tracaj" size="60" maxlength="60" value="<%= l_tradir %>">
						</td>
					</tr>
					<tr>
					    <td align="right"><b>Provincia:</b></td>
						<td >
							<input type="text" readonly class="deshabinp" name="transim" size="60" maxlength="60" value="<%= l_prodes %>">
						</td>
					</tr>
					<tr>
					    <td align="right"><b>Caja:</b></td>
						<td >
							<input type="text" readonly class="deshabinp" name="tracaj" size="60" maxlength="60" value="<%= l_tracaj %>">
						</td>
					</tr>
					<tr>
					    <td align="right"><b>Número Caja:</b></td>
						<td >
							<input type="text" readonly class="deshabinp" name="trancaj" size="60" maxlength="60" value="<%= l_tranrocaj %>">
						</td>
					</tr>
					<tr>
					    <td align="right"><b>Situación impositiva:</b></td>
						<td >
							<input type="text" readonly class="deshabinp" name="transim" size="60" maxlength="60" value="<%= l_sitimpdes %>">
						</td>
					</tr>
					</table>
				</td>
				<td width="5%"></td>
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
Cn.Close
Cn = nothing
%>
</body>
</html>

<% Option Explicit %>
<!--#include virtual="/serviciolocal/shared/db/conn_db.inc"-->
<% 
'Archivo: localidades_con_02.asp
'Descripción: Abm de Localidades
'Autor : Gustavo Manfrin
'Fecha: 08/02/2005
'Modificado: 

'Datos del formulario
Dim l_locnro
Dim l_loccod
Dim l_locdes
Dim l_pronro
Dim l_prodes

'ADO
Dim l_sql
Dim l_rs

%>
<html>
<head>
<link href="/serviciolocal/shared/css/tables3.css" rel="StyleSheet" type="text/css">
<meta http-equiv="Content-Type" content="text/html; charset=iso-8859-1">
<title><%= Session("Titulo")%>Localidades - Ticket</title>
</head>
<script src="/serviciolocal/shared/js/fn_ayuda.js"></script>
<script src="/serviciolocal/shared/js/fn_windows.js"></script>
<script>

</script>
<% 
Set l_rs = Server.CreateObject("ADODB.RecordSet")
l_locnro = request.querystring("cabnro")

l_sql = "SELECT locnro, loccod, locdes, tkt_localidad.pronro, prodes "
l_sql = l_sql & " FROM tkt_localidad  "
l_sql = l_sql & " INNER JOIN tkt_provincia ON tkt_provincia.pronro = tkt_localidad.pronro  "
l_sql = l_sql & " WHERE tkt_localidad.locnro = " & l_locnro

rsOpen l_rs, cn, l_sql, 0 
if not l_rs.eof then
	l_loccod = l_rs("loccod")
	l_locdes = l_rs("locdes")
	l_prodes = l_rs("prodes")
end if
l_rs.Close
%>
<body leftmargin="0" rightmargin="0" topmargin="0" bottommargin="0">
<form name="datos">

<table cellspacing="0" cellpadding="0" border="0" width="100%" height="100%">
<tr>
    <td class="th2"  nowrap>Localidades</td>
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
								<input type="text" readonly class="deshabinp" name="loccod" size="20" maxlength="15" value="<%= l_loccod %>">
							</td>
						</tr>
						<tr>
						    <td align="right"><b>Descripción:</b></td>
							<td>
								<input type="text" readonly class="deshabinp" name="locdes" size="60" maxlength="50" value="<%= l_locdes %>">
							</td>
						</tr>
						<tr>
						    <td align="right"><b>Provincia:</b></td>
							<td>
								<input type="text" readonly class="deshabinp" name="prodes" size="60" maxlength="60" value="<%= l_prodes %>">
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
Cn.Close
Cn = nothing
%>
</body>
</html>

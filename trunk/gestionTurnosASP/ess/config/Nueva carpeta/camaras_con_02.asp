<% Option Explicit %>
<!--#include virtual="/serviciolocal/shared/db/conn_db.inc"-->
<% 
'Archivo: camaras_con_02.asp
'Descripción: Abm de Cámaras
'Autor : Raul Chinestra
'Fecha: 08/02/2005

'Datos del formulario
Dim l_camnro
Dim l_camcod
Dim l_camdes

'ADO
Dim l_sql
Dim l_rs

%>
<html>
<head>
<link href="/serviciolocal/shared/css/tables3.css" rel="StyleSheet" type="text/css">
<meta http-equiv="Content-Type" content="text/html; charset=iso-8859-1">
<title><%= Session("Titulo")%>Cámaras - Ticket</title>
</head>
<script src="/serviciolocal/shared/js/fn_ayuda.js"></script>
<script src="/serviciolocal/shared/js/fn_windows.js"></script>
<script>

</script>
<% 
Set l_rs = Server.CreateObject("ADODB.RecordSet")
l_camnro = request.querystring("cabnro")
l_sql = "SELECT  camcod, camdes "
l_sql = l_sql & " FROM tkt_camara "
l_sql  = l_sql  & " WHERE camnro = " & l_camnro
rsOpen l_rs, cn, l_sql, 0 
if not l_rs.eof then
	l_camcod = l_rs("camcod")
	l_camdes = l_rs("camdes")
end if
l_rs.Close
%>
<body leftmargin="0" rightmargin="0" topmargin="0" bottommargin="0">
<form name="datos">

<table cellspacing="0" cellpadding="0" border="0" width="100%" height="100%">
<tr>
    <td class="th2"  nowrap>Cámaras</td>
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
								<input type="text" readonly class="deshabinp" name="camcod" size="12" maxlength="20" value="<%= l_camcod %>">
							</td>
						</tr>
						<tr>
						    <td align="right"><b>Descripción:</b></td>
							<td>
								<input type="text" readonly class="deshabinp" name="camdes" size="60" maxlength="50" value="<%= l_camdes %>">
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

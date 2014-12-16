<% Option Explicit %>
<!--#include virtual="/serviciolocal/shared/db/conn_db.inc"-->
<% 
'Archivo: corredores_con_02.asp
'Descripción: Abm de corredores
'Autor : Gustavo Manfrin
'Fecha: 09/02/2005
'Modificado: 

'Datos del formulario
Dim l_vencornro
Dim l_vencorcod
Dim l_vencordes
Dim l_vencorrazsoc
Dim l_venhab

'ADO
Dim l_sql
Dim l_rs

%>
<html>
<head>
<link href="/serviciolocal/shared/css/tables3.css" rel="StyleSheet" type="text/css">
<meta http-equiv="Content-Type" content="text/html; charset=iso-8859-1">
<title><%= Session("Titulo")%>Corredores - Ticket</title>
</head>
<script src="/serviciolocal/shared/js/fn_ayuda.js"></script>
<script src="/serviciolocal/shared/js/fn_windows.js"></script>
<script>

</script>
<% 
Set l_rs = Server.CreateObject("ADODB.RecordSet")
l_vencornro = request.querystring("cabnro")

l_sql = "SELECT vencornro, vencorcod, vencordes, vencorrazsoc, venhab "
l_sql = l_sql & " FROM tkt_vencor "
l_sql = l_sql & " WHERE venact = -1 AND vencortip = 'C' "
l_sql = l_sql & " AND vencornro = " & l_vencornro

rsOpen l_rs, cn, l_sql, 0 
if not l_rs.eof then
	l_vencornro = l_rs("vencornro")
	l_vencorcod = l_rs("vencorcod")
	l_vencordes = l_rs("vencordes")
	l_vencorrazsoc = l_rs("vencorrazsoc")
	l_venhab = l_rs("venhab")
end if
l_rs.Close
%>
<body leftmargin="0" rightmargin="0" topmargin="0" bottommargin="0">
<form name="datos">

<table cellspacing="0" cellpadding="0" border="0" width="100%" height="100%">
<tr>
    <td class="th2"  nowrap>Corredores</td>
	<td class="th2" align="right">
		<a class=sidebtnHLP href="Javascript:ayuda('<%= Request.ServerVariables("SCRIPT_NAME")%>');">Ayuda</a>
	</td>
</tr>
<tr>
	<td colspan="2" height="100%">
		<table width="100%" cellpadding="0" cellspacing="0" border="0">
			<tr>
				<td width="50%"></td>
				<td>
					<table cellpadding="0" cellspacing="0">
						<tr>
						    <td align="right"><b>Código:</b></td>
							<td>
								<input type="text" readonly class="deshabinp" name="corcod" size="20" maxlength="15" value="<%= l_vencorcod %>">
							</td>
						</tr>
						<tr>
						    <td align="right" nowrap><b>Nombre clave:</b></td>
							<td>
								<input type="text" readonly class="deshabinp" name="cordes" size="60" maxlength="50" value="<%= l_vencordes %>">
							</td>
						</tr>
						<tr>
						    <td align="right"  nowrap><b>Razón social:</b></td>
							<td>
								<input type="text" readonly class="deshabinp" name="corrzso" size="60" maxlength="50" value="<%= l_vencorrazsoc %>">
							</td>
						</tr>
						<tr>
						    <td align="right"  nowrap><b>Habilitado:</b></td>
							<td>
								<input type="Checkbox" readonly disabled name="corhab" <% If l_rs("venhab") = -1 then %>Checked<% end if %>>		
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

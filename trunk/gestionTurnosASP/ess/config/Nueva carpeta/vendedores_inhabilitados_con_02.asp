<% Option Explicit %>
<!--#include virtual="/ticket/shared/db/conn_db.inc"-->
<% 
'Archivo: vendedores_inhabilitados_con_02.asp
'Descripción: Abm de Vendedores inhabilitados
'Autor : Lisandro Moro
'Fecha: 09/02/2005

'Datos del formulario
'
on error goto 0

Dim l_vencornro
Dim l_vencordes
Dim l_vencorrazsoc
Dim l_venhab

'ADO
Dim l_sql
Dim l_rs

%>
<html>
<head>
<link href="/ticket/shared/css/tables3.css" rel="StyleSheet" type="text/css">
<meta http-equiv="Content-Type" content="text/html; charset=iso-8859-1">
<title><%= Session("Titulo")%>Vendedores/Corredores Inhabilitados - Ticket</title>
</head>
<script src="/ticket/shared/js/fn_ayuda.js"></script>
<script src="/ticket/shared/js/fn_windows.js"></script>
<script>

</script>
<% 
Set l_rs = Server.CreateObject("ADODB.RecordSet")
l_vencornro = request.querystring("cabnro")
l_sql = "SELECT vencornro, vencordes, vencorrazsoc, venhab"
l_sql = l_sql & " FROM tkt_vencor "
l_sql = l_sql & " WHERE vencornro = " & l_vencornro
rsOpen l_rs, cn, l_sql, 0 
if not l_rs.eof then
	l_vencornro = l_rs("vencornro")
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
    <td class="th2"  nowrap>Vendedores/Corredores Inhabilitados</td>
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
						    <td height="100%" align="right" nowrap><b>Descripción:</b></td>
							<td height="100%">
								<input type="text" readonly class="deshabinp" name="vencordes" size="50" maxlength="50" value="<%= l_vencordes %>">
							</td>
						</tr>
						<tr>
						    <td height="100%" align="right" nowrap><b>Razón Social:</b></td>
							<td height="100%">
								<input type="text" readonly class="deshabinp" name="vencorrazsoc" size="50" maxlength="50" value="<%= l_vencorrazsoc %>">
							</td>
						</tr>
						<tr>
						    <td height="100%" align="right" nowrap><b>Habilitado:</b></td>
							<td height="100%">
								<input type="checkbox" name="venhab" value="<%= l_venhab %>" <% If l_venhab = -1 then %>Checked<% End If %>>
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
		<% call MostrarBoton ("sidebtnABM", "Javascript:abrirVentanaH('vendedores_inhabilitados_con_03.asp?cabnro=" & l_vencornro & "&venhab=' + document.datos.venhab.checked,'',520,160);","Aceptar")%>
		<!--<a class=sidebtnABM href="Javascript:abrirVentanaH('vendedores_inhabilitados_con_03.asp?cabnro=<%= l_vencornro %>&venhab=' + document.datos.venhab.checked,'',520,160);">Aceptar</a>-->
		<a class=sidebtnABM href="Javascript:window.close()">Salir</a>
	</td>
</tr>
</table>
</form>
<%
set l_rs = nothing
Cn.Close
'Cn = nothing
%>
</body>
</html>

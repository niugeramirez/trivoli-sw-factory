<% Option Explicit %>
<!--#include virtual="/ticket/shared/db/conn_db.inc"-->
<% 
'Archivo: rubros_con_02.asp
'Descripción: Consulta de Rubros
'Autor : Raul Chinestra
'Fecha: 21/04/2005


'Datos del formulario
Dim l_rubnro
Dim l_rubdes
Dim l_rubcod
Dim l_rubabr

'ADO
Dim l_sql
Dim l_rs

%>
<html>
<head>
<link href="/ticket/shared/css/tables3.css" rel="StyleSheet" type="text/css">
<meta http-equiv="Content-Type" content="text/html; charset=iso-8859-1">
<title><%= Session("Titulo")%>Rubros - Ticket</title>
</head>
<script src="/ticket/shared/js/fn_ayuda.js"></script>
<script src="/ticket/shared/js/fn_windows.js"></script>
<script>

</script>
<% 
Set l_rs = Server.CreateObject("ADODB.RecordSet")
l_rubnro = request.querystring("cabnro")
l_sql = "SELECT  rubdes, rubcod, rubabr"
l_sql = l_sql & " FROM tkt_rubro "
l_sql  = l_sql  & " WHERE rubnro = " & l_rubnro
rsOpen l_rs, cn, l_sql, 0 
if not l_rs.eof then
	l_rubdes = l_rs("rubdes")
	l_rubcod = l_rs("rubcod")
	l_rubabr = l_rs("rubabr")
end if
l_rs.Close
%>
<body leftmargin="0" rightmargin="0" topmargin="0" bottommargin="0">
<form name="datos">

<table cellspacing="0" cellpadding="0" border="0" width="100%" height="100%">
<tr>
    <td class="th2"  nowrap>Rubros</td>
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
						    <td align="right" nowrap><b>Código:</b></td>
							<td>
								<input type="text" readonly class="deshabinp" name="rubcod" size="15" maxlength="12" value="<%= l_rubcod %>">
							</td>
						</tr>
						<tr>
						    <td align="right"><b>Descripción:</b></td>
							<td>
								<input type="text" readonly class="deshabinp" name="rubdes" size="50" maxlength="50" value="<%= l_rubdes %>">
							</td>
						</tr>
						<tr>
						    <td align="right"><b>Abreviatura:</b></td>
							<td>
								<input type="text" readonly class="deshabinp" name="rubabr" size="15" maxlength="40" value="<%= l_rubabr %>">
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

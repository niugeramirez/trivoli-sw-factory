<% Option Explicit %>
<!--#include virtual="/ticket/shared/db/conn_db.inc"-->
<% 
'Archivo: Tipos_documentos_con_02.asp
'Descripción: Abm de Tipos de Documentos
'Autor : Raul Chinestra
'Fecha: 08/02/2005

'Datos del formulario
Dim l_tipdocnro
Dim l_tipdocsig
Dim l_tipdocdes

'ADO
Dim l_sql
Dim l_rs

%>
<html>
<head>
<link href="/ticket/shared/css/tables3.css" rel="StyleSheet" type="text/css">
<meta http-equiv="Content-Type" content="text/html; charset=iso-8859-1">
<title><%= Session("Titulo")%>Tipos de Documentos - Ticket</title>
</head>
<script src="/ticket/shared/js/fn_ayuda.js"></script>
<script src="/ticket/shared/js/fn_windows.js"></script>
<script>

</script>
<% 
Set l_rs = Server.CreateObject("ADODB.RecordSet")
l_tipdocnro = request.querystring("cabnro")
l_sql = "SELECT  tipdocsig, tipdocdes "
l_sql = l_sql & " FROM tkt_tipodocumento "
l_sql  = l_sql  & " WHERE tipdocnro = " & l_tipdocnro
rsOpen l_rs, cn, l_sql, 0 
if not l_rs.eof then
	l_tipdocsig = l_rs("tipdocsig")
	l_tipdocdes = l_rs("tipdocdes")
end if
l_rs.Close
%>
<body leftmargin="0" rightmargin="0" topmargin="0" bottommargin="0">
<form name="datos">

<table cellspacing="0" cellpadding="0" border="0" width="100%" height="100%">
<tr>
    <td class="th2"  nowrap>Tipos de Documentos</td>
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
						    <td height="50%" align="right"><b>Sigla:</b></td>
							<td height="50%">
								<input type="text" readonly class="deshabinp" name="tipdocsig" size="60" maxlength="50" value="<%= l_tipdocsig %>">
							</td>
						</tr>
						<tr>
						    <td height="50%" align="right"><b>Descripción</b></td>
							<td height="50%">
								<input type="text" readonly class="deshabinp" name="tipdocdes" size="60" maxlength="50" value="<%= l_tipdocdes %>">
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

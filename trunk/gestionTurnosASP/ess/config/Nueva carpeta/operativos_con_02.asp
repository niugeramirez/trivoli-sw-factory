<% Option Explicit %>
<!--#include virtual="/ticket/shared/db/conn_db.inc"-->
<% 
'Archivo: operativos_con_02.asp
'Descripción: Abm de Cámaras
'Autor : Lisandro Moro
'Fecha: 09/02/2005

'Datos del formulario
'pronro, procod, prodes, tippronro, proenv, procla, provercon, promez
on error goto 0

Dim l_openro

Dim l_opecod
Dim l_opecan
Dim l_lugdes
Dim l_opefeclle
Dim l_opehorlle
Dim l_opetip

Dim l_Transporte
'ADO
Dim l_sql
Dim l_rs

%>
<html>
<head>
<link href="/ticket/shared/css/tables3.css" rel="StyleSheet" type="text/css">
<meta http-equiv="Content-Type" content="text/html; charset=iso-8859-1">
<title><%= Session("Titulo")%>Carga de Operativos - Ticket</title>
</head>
<style type="text/css">
.none{
	padding : 0;
	padding-left : 0;
}
</style>
<script src="/ticket/shared/js/fn_ayuda.js"></script>
<script src="/ticket/shared/js/fn_windows.js"></script>
<script>

</script>
<% 
Set l_rs = Server.CreateObject("ADODB.RecordSet")
l_openro = request.querystring("cabnro")
l_sql = "SELECT opecod, opecan, lugdes, opefeclle, opehorlle, opetip "
l_sql = l_sql & " FROM tkt_operativo "
l_sql = l_sql & " INNER JOIN tkt_lugar ON tkt_lugar.lugnro = tkt_operativo.lugnro "
l_sql  = l_sql  & " WHERE openro = " & l_openro
rsOpen l_rs, cn, l_sql, 0 
if not l_rs.eof then
	l_opecod = l_rs("opecod")
	l_opecan = l_rs("opecan")
	l_lugdes = l_rs("lugdes")
	l_opefeclle = l_rs("opefeclle")
	l_opehorlle = left(l_rs("opehorlle"),2) &":"& right(l_rs("opehorlle"),2)
	l_opetip = l_rs("opetip")
end if
l_rs.Close
%>
<body leftmargin="0" rightmargin="0" topmargin="0" bottommargin="0">
<form name="datos">

<table cellspacing="0" cellpadding="0" border="0" width="100%" height="100%">
<tr>
    <td class="th2"  nowrap>Carga de Operativos</td>
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
						    <td height="100%" align="right" nowrap><b>Código:</b></td>
							<td height="100%" class="none">
								<table cellpadding="0" cellspacing="0" border="0">
									<tr>
										<td>
											<input type="text" readonly class="deshabinp" name="opecod" size="12" value="<%= l_opecod %>">
										</td>
									    <td width="100%" align="right" nowrap><b>Cantidad:</b></td>
										<td>
											<input type="text" readonly class="deshabinp" name="opecan" size="10" maxlength="50" value="<%= l_opecan %>">
										</td>
									</tr>
								</table>
							</td>
						</tr>
						<tr>
						    <td height="100%" align="right" nowrap><b>Procedencia:</b></td>
							<td height="100%">
								<input type="text" readonly class="deshabinp" name="lugdes" size="50" maxlength="50" value="<%= l_lugdes %>">
							</td>
						</tr>
						<tr>
						    <td height="100%" align="right" nowrap><b>Fecha Llegada:</b></td>
							<td height="100%" class="none">
								<table cellpadding="0" cellspacing="0" border="0">
									<tr>
										<td>
											<input type="text" class="deshabinp"  readonly  name="opefeclle" size="10" value="<%= l_opefeclle %>">
										</td>
									    <td width="100%" align="right" nowrap><b>Hora Llegada:</b></td>
										<td>
											<input type="text" class="deshabinp" readonly  name="opehorlle" size="10" value="<%= l_opehorlle %>">
										</td>
									</tr>
								</table>
							</td>
						</tr>
						<tr>
						    <td height="100%" align="right" nowrap><b>Tipo:</b></td>
							<td height="100%">
								<input type="radio" disabled name="opetip" <% if l_opetip = "C" then%>checked<% end if%>>Carga 
								<input type="radio" disabled name="opetip" <% if l_opetip = "D" then%>checked<% end if%>>Descarga								
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
'Cn = nothing
%>
</body>
</html>

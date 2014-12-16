<% Option Explicit %>
<!--#include virtual="/ticket/shared/db/conn_db.inc"-->
<% 
'Archivo: ordenes_con_02.asp
'Descripción: Consulta de Ordenes de trabajo
'Autor : Alvaro Bayon
'Fecha: 09/02/2005
'Modificado: 

'Datos del formulario
Dim l_ordnro
Dim l_prodes
Dim l_ordkil
Dim l_orilugdes
Dim l_deslugdes
Dim l_ordhab
Dim l_ordcod
Dim l_ordobs
Dim l_trades

'ADO
Dim l_sql
Dim l_rs

%>
<html>
<head>
<link href="/ticket/shared/css/tables3.css" rel="StyleSheet" type="text/css">
<meta http-equiv="Content-Type" content="text/html; charset=iso-8859-1">
<title><%= Session("Titulo")%>Ordenes de trabajo - Ticket</title>
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
l_ordnro = request.querystring("cabnro")
l_sql = "SELECT ordnro, ordcod, prodes, tkt_lugar.lugdes, ordhab, lugard.lugdes as deslugdes, ordkil, ordobs "', trades "
l_sql = l_sql & " FROM tkt_ordentrabajo "
l_sql = l_sql & " INNER JOIN tkt_producto ON tkt_ordentrabajo.pronro = tkt_producto.pronro"
'l_sql = l_sql & " INNER JOIN tkt_transportista ON tkt_ordentrabajo.tranro = tkt_transportista.tranro"
l_sql = l_sql & " INNER JOIN tkt_lugar ON tkt_ordentrabajo.orilugnro = tkt_lugar.lugnro"
l_sql = l_sql & " INNER JOIN tkt_lugar lugard  ON tkt_ordentrabajo.deslugnro = lugard.lugnro"
l_sql  = l_sql  & " WHERE ordnro = " & l_ordnro
rsOpen l_rs, cn, l_sql, 0 
if not l_rs.eof then
	l_prodes = l_rs("prodes")
	l_ordkil = l_rs("ordkil")
	l_orilugdes = l_rs("lugdes")
	l_deslugdes = l_rs("deslugdes")
	l_ordhab = l_rs("ordhab")
	l_ordcod = l_rs("ordcod")
	l_ordobs = l_rs("ordobs")
	l_trades = l_rs("trades")
end if
l_rs.Close
%>
<body leftmargin="0" rightmargin="0" topmargin="0" bottommargin="0">
<form name="datos">

<table cellspacing="0" cellpadding="0" border="0" width="100%" height="100%">
<tr>
    <td class="th2"  nowrap>Ordenes de trabajo</td>
	<td class="th2" align="right">
		<a class=sidebtnHLP href="Javascript:ayuda('<%= Request.ServerVariables("SCRIPT_NAME")%>');">Ayuda</a>
	</td>
</tr>
<tr>
	<td colspan="2" height="100%">
		<table border="0" cellspacing="0" cellpadding="0">
			<tr>
				<td width="10%"></td>
				<td>
					<table cellspacing="0" cellpadding="0" border="0">
					<tr>
					    <td align="right"><b>Código:</b></td>
						<td>
							<input type="text" readonly class="deshabinp" name="ordcod" size="12" maxlength="10" value="<%= l_ordcod %>">
						</td>
					</tr>
					<tr>
					    <td align="right"><b>Producto:</b></td>
						<td>
							<input type="text" readonly class="deshabinp" name="prodes" size="60" maxlength="50" value="<%= l_prodes %>">
						</td>
					</tr>
					<tr>
					    <td align="right"><b>Origen:</b></td>
						<td>
							<input type="text" readonly class="deshabinp" name="orilugdes" size="60" maxlength="50" value="<%= l_orilugdes %>">
						</td>
					</tr>
					<tr>
					    <td align="right"><b>Destino:</b></td>
						<td>
							<input type="text" readonly class="deshabinp" name="deslugdes" size="60" maxlength="50" value="<%= l_deslugdes %>">
						</td>
					</tr>
					<tr>
					    <td align="right"><b>Kilos estimados:</b></td>
						<td>
							<input type="text" readonly class="deshabinp" name="ordkil" size="60" maxlength="50" value="<%= l_ordkil %>">
						</td>
					</tr>
					<tr>
					    <td align="right"><b>Habilitada:</b></td>
						<td>
							<input type="Checkbox" readonly disabled  name="ordhab" <% If l_ordhab = -1 then %>Checked<% end if %>>				
						</td>
					</tr>
					<tr>
					    <td align="right"><b>Observaciones:</b></td>
						<td>
							<TEXTAREA name="ordobs" readonly class="deshabinp" rows="4" cols="45" ><%= l_ordobs %></TEXTAREA>
						</td>
					</tr>
					<tr><td colspan="2"><br></td></tr>
					<tr>
						<td colspan="2" class="none">
							<table cellpadding="0" cellspacing="0" width="100%" height="100%" border="1">
								<tr>
									<td width="50%" class="barra">Transportistas</td>
									<td width="50%" class="barra">Camioneros</td>
								</tr>
								<tr>
									<td class="none">
										<iframe name="trans" src="ordenes_transportistas_con_01.asp?ordnro=<%= l_ordnro %>"></iframe>
									</td>
									<td  class="none">
										<iframe name="camiones" src="ordenes_camioneros_con_01.asp?ordnro=0&tranro=0"></iframe>
									</td>
								</tr>
							</table>
						</td>
					</tr>
					</table>
				</td>
				<td width="10%"></td>
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

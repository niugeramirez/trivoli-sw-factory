<% Option Explicit %>
<!--#include virtual="/serviciolocal/shared/db/conn_db.inc"-->
<% 
'Archivo: contratos_con_02.asp
'Descripción: Abm de Contratos
'Autor : Lisandro Moro
'Fecha: 11/02/2005

'Datos del formulario
on error goto 0

Dim l_connro
Dim l_concod
Dim l_confec
Dim l_pronro
Dim l_empnro
Dim l_vennro
Dim l_desnro
Dim l_entnro
Dim l_cornro
Dim l_tipopenro
Dim l_tipconnro
Dim l_contip
Dim l_lugdes
Dim l_conact
Dim l_conobs
Dim l_conkil
Dim l_conkilent
Dim l_conkilsal
Dim l_conprodes

Dim l_prodes
Dim l_vencordes
Dim l_desdes
Dim l_cordes
dim l_empdes
Dim l_vendes
Dim l_entdes
Dim l_tipopedes
Dim l_condes
Dim l_conhas

'ADO
Dim l_sql
Dim l_rs

%>
<html>
<head>
<link href="/serviciolocal/shared/css/tables3.css" rel="StyleSheet" type="text/css">
<meta http-equiv="Content-Type" content="text/html; charset=iso-8859-1">
<title><%= Session("Titulo")%>Contratos - Ticket</title>
</head>
<style type="text/css">
.none{
	padding : 0;
	padding-left : 0;
}
</style>
<script src="/serviciolocal/shared/js/fn_ayuda.js"></script>
<script src="/serviciolocal/shared/js/fn_windows.js"></script>
<script>

</script>
<% 
Set l_rs = Server.CreateObject("ADODB.RecordSet")
l_connro = request.querystring("cabnro")



l_sql = " SELECT vendedor.vencordes as vendes, corredor.vencordes as cordes, "
l_sql = l_sql & " connro, concod, confec, prodes, empdes, desdes, entdes, tipopedes, "
l_sql = l_sql & " tipconnro, contip, lugdes, conact, conobs, conkil, conkilent, conkilsal "
l_sql = l_sql & " , conprodes, condesde, conhasta "
l_sql = l_sql & " FROM tkt_contrato "
l_sql = l_sql & " LEFT JOIN tkt_producto ON tkt_producto.pronro = tkt_contrato.pronro "
l_sql = l_sql & " LEFT JOIN tkt_empresa ON tkt_empresa.empnro = tkt_contrato.empnro "
l_sql = l_sql & " LEFT JOIN tkt_vencor vendedor ON vendedor.vencornro = tkt_contrato.vennro "
l_sql = l_sql & " LEFT JOIN tkt_vencor corredor ON corredor.vencornro = tkt_contrato.cornro "
l_sql = l_sql & " LEFT JOIN tkt_destinatario ON tkt_destinatario.desnro = tkt_contrato.desnro "
l_sql = l_sql & " LEFT JOIN tkt_entrec ON tkt_entrec.entnro = tkt_contrato.entnro "
l_sql = l_sql & " LEFT JOIN tkt_tipooperacion ON tkt_tipooperacion.tipopenro = tkt_contrato.tipopenro "
l_sql = l_sql & " LEFT JOIN tkt_lugar ON tkt_lugar.lugnro = tkt_contrato.lugnro "
l_sql  = l_sql  & " WHERE connro = " & l_connro
rsOpen l_rs, cn, l_sql, 0 
if not l_rs.eof then
	l_connro = l_rs("connro")
	l_concod = l_rs("concod")
	l_confec = l_rs("confec")
	'l_pronro = l_rs("pronro")
	'l_empnro = l_rs("empnro")
	'l_vennro = l_rs("vencornro")
	'l_desnro = l_rs("desnro")
	l_entdes = l_rs("entdes")
	'l_cordes = l_rs("cordes")
	l_tipopedes = l_rs("tipopedes")
	l_tipconnro = l_rs("tipconnro")
	l_contip = l_rs("contip")
	l_lugdes = l_rs("lugdes")
	l_conact = l_rs("conact")
	l_conobs = l_rs("conobs")
	l_conkil = l_rs("conkil")
	l_conkilent = l_rs("conkilent")
	l_conkilsal = l_rs("conkilsal")
	l_vencordes = l_rs("vendes")
	l_prodes = l_rs("prodes")
	l_desdes = l_rs("desdes")
	l_cordes = l_rs("cordes")
	l_empdes = l_rs("empdes")
	l_vendes = l_rs("vendes")
	l_cordes = l_rs("cordes")
	l_conprodes = l_rs("conprodes")
	l_condes = l_rs("condesde")
	l_conhas = l_rs("conhasta")
	
end if
l_rs.Close
%>
<body leftmargin="0" rightmargin="0" topmargin="0" bottommargin="0">
<form name="datos">

<table cellspacing="0" cellpadding="0" border="0" width="100%" height="100%">
<tr>
    <td class="th2"  nowrap>Contratos</td>
	<td class="th2" align="right">
		<a class=sidebtnHLP href="Javascript:ayuda('<%= Request.ServerVariables("SCRIPT_NAME")%>');">Ayuda</a>
	</td>
</tr>
<tr>
	<td colspan="2"  height="100%">
		<table border="0" cellspacing="0" cellpadding="0">
			<tr>
				<td width="50%"></td>
				<td>
					<table cellspacing="0" cellpadding="0" border="0">
						<tr>
						    <td align="right" nowrap><b>Código:</b></td>
							<td>
								<table border="0" cellpadding="0" cellspacing="0" width="100%">
									<tr>
										<td align="left" class="none">
											<input type="text" readonly class="deshabinp" name="concod" size="12" maxlength="12" value="<%= l_concod %>">
										</td>
										<td width="100%" nowrap align="right"><b>Fecha:&nbsp;</b></td>
										<td class="none">
											<input type="text" readonly class="deshabinp" name="confec" size="10" maxlength="10" value="<%= l_confec %>">
										</td>
									</tr>
								</table>
							</td>
						</tr>
						<tr>
						    <td align="right" nowrap><b>Producto:</b></td>
							<td>
								<input type="text" readonly class="deshabinp" name="prodes" size="50" maxlength="50" value="<%= l_prodes %>">
							</td>
						</tr>
						<tr>
						    <td align="right" nowrap><b>Vendedor:</b></td>
							<td>
								<input type="text" readonly class="deshabinp" name="vencordes" size="50" maxlength="50" value="<%= l_vencordes %>">
							</td>
						</tr>
						<tr>
						    <td align="right" nowrap><b>Destinatario:</b></td>
							<td>
								<input type="text" readonly class="deshabinp" name="desdes" size="50" maxlength="50" value="<%= l_desdes %>">
							</td>
						</tr>
						<tr>
						    <td align="right" nowrap><b>Entregador:</b></td>
							<td>
								<input type="text" readonly class="deshabinp" name="entdes" size="50" maxlength="50" value="<%= l_entdes %>">
							</td>
						</tr>
						<tr>
						    <td align="right" nowrap><b>Corredor:</b></td>
							<td>
								<input type="text" readonly class="deshabinp" name="cordes" size="50" maxlength="50" value="<%= l_cordes %>">
							</td>
						</tr>
						<tr>
						    <td align="right" nowrap><b>Empresa:</b></td>
							<td>
								<input type="text" readonly class="deshabinp" name="empdes" size="50" maxlength="50" value="<%= l_empdes %>">
							</td>
						</tr>
						<tr>
						    <td align="right" nowrap><b>Tipo Operación:</b></td>
							<td>
								<table border="0" cellpadding="0" cellspacing="0" width="100%">
									<tr>
										<td align="left" class="none">
											<input type="text" readonly class="deshabinp" name="tipopedes" size="10" maxlength="10" value="<%= l_tipopedes %>">
										</td>
										<td width="100%" nowrap align="right"><b>Kilos Totales:&nbsp;</b></td>
										<td class="none">
											<input type="text" readonly class="deshabinp" name="conkil" size="10" maxlength="10" value="<%= l_conkil %>">
										</td>
									</tr>
								</table>
							</td>
						</tr>
						<tr>
						    <td align="right" nowrap><b>Tipo Contrato:</b></td>
							<td>
								<table border="0" cellpadding="0" cellspacing="0" width="100%">
									<tr>
										<td align="left" class="none">
											<input type="text" readonly class="deshabinp" name="tipcon" size="10" maxlength="10" value="<%= l_tipconnro %>">
										</td>
										<td width="100%" nowrap align="right"><b>Kilos Entregados:&nbsp;</b></td>
										<td class="none">
											<input type="text" readonly class="deshabinp" name="conkilent" size="10" maxlength="10" value="<%= l_conkilent %>">
										</td>
									</tr>
								</table>
							</td>
						</tr>
						<tr>
						    <td align="right" nowrap><b>Granel-Bolsa:</b></td>
							<td>
								<table border="0" cellpadding="0" cellspacing="0" width="100%">
									<tr>
										<td align="left" class="none">
											<input type="text" readonly class="deshabinp" name="contip" size="10" maxlength="10" value="<%if l_contip="G" then %>Granel<% Else %>Bolsa<% End If %>">
										</td>
										<td width="100%" nowrap align="right"><b>Kilos Saldo:&nbsp;</b></td>
										<td class="none">
											<input type="text" readonly class="deshabinp" name="conkilsal" size="10" maxlength="10" value="<%= l_conkilsal %>">
										</td>
									</tr>
								</table>
							</td>
						</tr>
   				        <tr>
						    <td align="right" nowrap><b>Entrega Desde:</b></td>
							<td>
								<table border="0" cellpadding="0" cellspacing="0" width="100%">
									<tr>
										<td align="left" class="none">
											<input type="text" readonly class="deshabinp" name="condes" size="10" maxlength="10" value="<%=l_condes%>">
										</td>
										<td width="100%" nowrap align="right"><b>Entrega Hasta:&nbsp;</b></td>
										<td class="none">
											<input type="text" readonly class="deshabinp" name="conhas" size="10" maxlength="10" value="<%= l_conhas %>">
										</td>
									</tr>
								</table>
							</td>
						</tr>
						<tr>
						    <td align="right" nowrap><b>Lugar Entrega:</b></td>
							<td>
								<input type="text" readonly class="deshabinp" name="lugdes" size="50" maxlength="50" value="<%= l_lugdes %>">
							</td>
						</tr>
						<tr>
						    <td align="right" nowrap><b>Activa:</b></td>
							<td>
								<table border="0" cellpadding="0" cellspacing="0" width="100%">
									<tr>
										<td align="left" class="none">
											<input type="checkbox" readonly disabled name="conact" <% if l_conact = -1 then %>checked<% End If %>>
										</td>
										<td width="100%" nowrap align="right"><b>Kilos (Destino - Procedencia):&nbsp;</b></td>
										<td class="none">
											<input type="text" readonly class="deshabinp" name="conprodes" size="10" maxlength="10" value="<%= l_conprodes %>">
										</td>
									</tr>
								</table>

							</td>
						</tr>
						<tr>
						    <td align="right" nowrap><b>Observaciones:</b></td>
							<td>
								<!--<input type="text" readonly class="deshabinp" name="conobs" size="50" maxlength="10"  value="<%'= l_conobs %>">-->
								<textarea cols="50" style="width:325px" class="deshabinp" readonly rows="3" name="conobs" ><%= l_conobs %></textarea>
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

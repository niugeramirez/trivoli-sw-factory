<% Option Explicit %>
<!--#include virtual="/ticket/shared/db/conn_db.inc"-->
<% 
'Archivo: param_configuracion_con_01.asp
'Descripción: Abm de parametros de configuracion
'Autor : Lisandro Moro
'Fecha: 01/03/2005
' Modificado: Raul CHinestra - 15/06/2006 - Se agregó la cantidad de dias que se desea mostrar el transito

'on error goto 0
Dim l_rs
Dim l_sql
Dim l_filtro
Dim l_orden
Dim l_sqlfiltro
Dim l_sqlorden
Dim l_todos

Dim l_configura
Dim l_productiv
Dim l_habilita
Dim l_casrem
Dim l_clapes
Dim l_merxvag
Dim l_difpesxven
Dim l_operati
Dim l_proxlin
Dim l_bloconcum
Dim l_nroins
Dim l_netdpcam
Dim l_netdpvag
Dim l_brudpcam
Dim l_brudpvag
Dim l_summez
Dim l_brumaxcam
Dim l_ticpre
Dim l_rempre
Dim l_cantic
Dim l_canrem
Dim l_etiosob
Dim l_mostra
Dim l_cirpla

Set l_rs = Server.CreateObject("ADODB.RecordSet")

l_sql = "SELECT  configura, productiv, habilita, casrem, clapes, merxvag, difpesxven, operati, proxlin, "
l_sql  = l_sql  & " bloconcum, nroins, netdpcam, netdpvag, brudpcam, brudpvag, summez, brumaxcam, ticpre, "
l_sql  = l_sql  & "  rempre, cantic, canrem, etiosob, mostra, cirpla "
l_sql  = l_sql  & " FROM tkt_config "
rsOpen l_rs, cn, l_sql, 0
if not l_rs.eof then
	l_configura	= l_rs("configura")
	l_productiv = l_rs("productiv")
	l_habilita = l_rs("habilita")
	l_casrem = l_rs("casrem")
	l_clapes = l_rs("clapes")
	l_merxvag = l_rs("merxvag")
	l_difpesxven = l_rs("difpesxven")
	l_operati = l_rs("operati")
	l_proxlin = l_rs("proxlin")
	l_bloconcum = l_rs("bloconcum")
	l_nroins = l_rs("nroins")
	l_netdpcam = l_rs("netdpcam")
	l_netdpvag = l_rs("netdpvag")
	l_brudpcam = l_rs("brudpcam")
	l_brudpvag = l_rs("brudpvag")
	l_summez = l_rs("summez")
	l_brumaxcam = l_rs("brumaxcam")
	l_ticpre = l_rs("ticpre")
	l_rempre = l_rs("rempre")
	l_cantic = l_rs("cantic")
	l_canrem = l_rs("canrem")
	l_etiosob = l_rs("etiosob")
	l_cirpla = l_rs("cirpla")
	l_mostra = l_rs("mostra")
else
	l_configura	= ""
	l_productiv = 0
	l_habilita = 0
	l_casrem = ""
	l_clapes = 0
	l_merxvag = 0
	l_difpesxven = 0
	l_operati = 0
	l_proxlin = 0
	l_bloconcum = 0
	l_nroins = 0
	l_netdpcam = 0
	l_netdpvag = 0
	l_brudpcam = 0
	l_brudpvag = 0
	l_summez = 0
	l_brumaxcam = 0
	l_ticpre = 0
	l_rempre = 0
	l_cantic = 0
	l_canrem = 0
	l_etiosob = ""
	l_cirpla = ""
	l_mostra = ""
end if
l_rs.close
%>
<!DOCTYPE HTML PUBLIC "-//IETF//DTD HTML//EN">
<html>
<script src="/ticket/shared/js/fn_windows.js"></script>
<script src="/ticket/shared/js/fn_confirm.js"></script>
<script src="/ticket/shared/js/fn_ayuda.js"></script>
<script src="/ticket/shared/js/fn_valida.js"></script>
<head>
<link href="/ticket/shared/css/tables3.css" rel="StyleSheet" type="text/css">
<meta http-equiv="Content-Type" content="text/html; charset=iso-8859-1">
<title><%= Session("Titulo")%>Parámetros de Configuración - Ticket</title>
</head>

<script>
function Valida(){
	document.datos.nroins.value = Trim(document.datos.nroins.value);
	document.datos.netdpcam.value = Trim(document.datos.netdpcam.value);
	document.datos.netdpvag.value = Trim(document.datos.netdpvag.value);
	document.datos.brudpcam.value = Trim(document.datos.brudpcam.value);
	document.datos.brudpvag.value = Trim(document.datos.brudpvag.value);
	document.datos.summez.value = Trim(document.datos.summez.value);
	document.datos.brumaxcam.value = Trim(document.datos.brumaxcam.value);
	document.datos.cantic.value = Trim(document.datos.cantic.value);
	document.datos.canrem.value = Trim(document.datos.canrem.value);
	document.datos.mostra.value = Trim(document.datos.mostra.value);	
	if (isNaN(document.datos.nroins.value)){
		document.datos.nroins.select();
		alert('Debe ingresar un valor numérico \n en Número de Instalación.');
		document.datos.nroins.focus();
		return;
	}
	if (isNaN(document.datos.netdpcam.value)){
		document.datos.netdpcam.select();
		alert('Debe ingresar un valor numérico \n en Neto D/P Camión.');
		document.datos.netdpcam.focus();
		return;
	}
	if (isNaN(document.datos.netdpvag.value)){
		document.datos.netdpvag.select();
		alert('Debe ingresar un valor numérico \n en Neto D/P Vagón.');
		document.datos.netdpvag.focus();
		return;
	}
	if (isNaN(document.datos.brudpcam.value)){
		document.datos.brudpcam.select();
		alert('Debe ingresar un valor numérico \n en Bruto D/P Camión.');
		document.datos.brudpcam.focus();
		return;
	}
	if (isNaN(document.datos.brudpvag.value)){
		document.datos.brudpvag.select();
		alert('Debe ingresar un valor numérico \n en Bruto D/P Vagón.');
		document.datos.brudpvag.focus();
		return;
	}	
	if (isNaN(document.datos.summez.value)){
		document.datos.summez.select();
		alert('Debe ingresar un valor numérico \n en Sumatoria Mezcla.');
		document.datos.summez.focus();
		return;
	}	
	if (isNaN(document.datos.brumaxcam.value)){
		document.datos.brumaxcam.select();
		alert('Debe ingresar un valor numérico \n en Bruto Maximo por Camión.');
		document.datos.brumaxcam.focus();
		return;
	}	
	if (isNaN(document.datos.cantic.value)){
		document.datos.cantic.select();
		alert('Debe ingresar un valor numérico \n en Cantidad de Tickets.');
		document.datos.cantic.focus();
		return;
	}	
	if (isNaN(document.datos.canrem.value)){
		document.datos.canrem.select();
		alert('Debe ingresar un valor numérico \n en Cantidad de Remitos.');
		document.datos.canrem.focus();
		return;
	}
	if (document.datos.mostra.value == "" ){
		document.datos.mostra.select();
		alert('Debe ingresar un valor en la Cantidad de días a mostrar Tránsitos.');
		document.datos.mostra.focus();
		return;
	}
	if (document.datos.mostra.value <= "0"){
		document.datos.mostra.select();
		alert('Debe ingresar un valor mayor a 0 en la Cantidad de días a mostrar Tránsitos.');
		document.datos.mostra.focus();
		return;
	}	
	if (document.datos.mostra.value >= "999"){
		document.datos.mostra.select();
		alert('Debe ingresar un valor menor a 999 en la Cantidad de días a mostrar Tránsitos.');
		document.datos.mostra.focus();
		return;
	}		
	if (isNaN(document.datos.mostra.value)){
		document.datos.mostra.select();
		alert('Debe ingresar un valor numérico \n en la Cantidad de días a mostrar Tránsitos.');
		document.datos.mostra.focus();
		return;
	}
	
	//abrirVentana('',"newa",500,500);
	//document.datos.target = "newa";
	document.datos.submit();
}
</script>

<body leftmargin="0" rightmargin="0" topmargin="0" bottommargin="0">
<form name="datos" action="param_configuracion_con_03.asp" method="post" target="ifrm2">
<table cellspacing="0" cellpadding="0" border="0" width="100%" height="100%">
<tr>
	<td colspan="2" height="100%">
		<table border="0" cellspacing="0" cellpadding="0">
			<tr>
				<td width="50%"></td>
				<td>
					<table cellspacing="0" cellpadding="0" border="0">
						<tr>
						    <td align="right" nowrap><b>Configuración:</b></td>
							<td colspan="3">
								<select name="configura" size="1" style="width:200;" <%'= l_claseCombo %>>
									<option value="" selected>&laquo; Seleccione una Configuración &raquo;</option>
									<option value="A">Acopio (A)</option>
									<option value="E">Entrega (E)</option>
									<option value="F">Fábrica (F)</option>
									<option value="G">Puerto (G)</option>
									<option value="P">Playa (P)</option>
									<option value="S">Servicios (S)</option>
								</select>
								<% 	If l_configura = "0" or l_configura = ""  or l_configura = " " or IsNull(l_configura) then
										l_configura = ""
									end if %>
									<script> document.datos.configura.value= "<%= l_configura %>"</script>
							</td>
						</tr>
						<tr>
						    <td align="right" nowrap><b>Productividad por Turnos:</b></td>
							<td>
								<input type="checkbox" disabled readonly class="deshabinp" name="productiv" <% If l_productiv = -1 then %>Checked<% End If %>>
							</td>
						    <td align="right" nowrap><b>Habilitación del Supervisor:</b></td>
							<td>
								<input type="checkbox" disabled readonly class="deshabinp" name="habilita" <% If l_habilita = -1 then %>Checked<% End If %>>
							</td>
						</tr>
						<tr>
						    <td align="right" nowrap><b>Remito por Cáscara:</b></td>
							<td colspan="3">
								<select name="casrem" disabled readonly class="deshabinp" size="1" style="width:200;" <%'= l_claseCombo %>>
									<option value="M">Por Movimiento</option>
									<option value="D">Por Día</option>
								</select>
								<% 	If l_casrem = "" or IsNull(l_casrem) then
										l_casrem = "0"
									end if %>
									<script> document.datos.casrem.value= "<%= l_casrem %>";</script>
							</td>
						</tr>
						<tr>
						    <td align="right" nowrap><b>Clave para Forzar Pesos:</b></td>
							<td>
								<input type="checkbox" name="clapes" <% If l_clapes = -1 then %>Checked<% End If %>>
							</td>
						    <td align="right" nowrap><b>Mercadería por Vagón:</b></td>
							<td>
								<input type="checkbox" disabled readonly class="deshabinp" name="merxvag" <% If l_merxvag = -1 then %>Checked<% End If %>>
							</td>
						</tr>
						<tr>
   						    <td align="right" nowrap><b>Diferencia Peso por Vendedor</b></td>
							<td>
								<input type="checkbox" disabled readonly class="deshabinp" name="difpesxven" <% If l_difpesxven = -1 then %>Checked<% End If %>>
							</td>
						    <td align="right" nowrap><b>Controla los Operativos:</b></td>
							<td>
								<input type="checkbox" disabled readonly class="deshabinp" name="operati" <% If l_operati = -1 then %>Checked<% End If %>>
							</td>
						</tr>
						<tr>
						    <td align="right" nowrap><b>Producto por Linea:</b></td>
							<td>
								<input type="checkbox" disabled readonly class="deshabinp" name="proxlin" <% If l_proxlin = -1 then %>Checked<% End If %>>
							</td>
						    <td align="right" nowrap><b>Bloquea Contrato Cumplido:</b></td>
							<td>
								<input type="checkbox" disabled readonly class="deshabinp" name="bloconcum" <% If l_bloconcum = -1 then %>Checked<% End If %>>
							</td>
						</tr>
						<tr>
						    <td align="right" nowrap><b>Número de Instalación:</b></td>
							<td colspan="3">
								<input type="text" name="nroins" size="5" maxlength="4" value="<%= l_nroins %>">
							</td>
						</tr>
						<tr>
						    <td align="right" nowrap><b>Neto D/P Camión:</b></td>
							<td>
								<input type="text" name="netdpcam" size="5" maxlength="5" value="<%= l_netdpcam %>">
							</td>
						    <td align="right" nowrap><b>Neto D/P Vagón:</b></td>
							<td>
								<input type="text" name="netdpvag" size="5" maxlength="5" value="<%= l_netdpvag %>">
							</td>
						</tr>
						<tr>
						    <td align="right" nowrap><b>Bruto D/P Camión:</b></td>
							<td>
								<input type="text" name="brudpcam" size="5" maxlength="5" value="<%= l_brudpcam %>">
							</td>
						    <td align="right" nowrap><b>Bruto D/P Vagón:</b></td>
							<td>
								<input type="text" name="brudpvag" size="5" maxlength="5" value="<%= l_brudpvag %>">
							</td>
						</tr>
						<tr>
						    <td align="right" nowrap><b>Sumatoria Mezcla:</b></td>
							<td>
								<input type="text" disabled readonly class="deshabinp" name="summez" size="5" maxlength="5" value="<%= l_summez %>">
							</td>
						    <td align="right" nowrap><b>Bruto Maximo  por Camión:</b></td>
							<td>
								<input type="text" name="brumaxcam" size="5" maxlength="5" value="<%= l_brumaxcam %>">
							</td>
						</tr>
						<tr>
						    <td align="right" nowrap><b>Ticket PreImpreso:</b></td>
							<td>
								<input type="checkbox" disabled readonly class="deshabinp" name="ticpre" <% If l_ticpre = -1 then %>Checked<% End If %>>
							</td>
						    <td align="right" nowrap><b>Remito PreImpreso:</b></td>
							<td>
								<input type="checkbox" name="rempre" <% If l_rempre = -1 then %>Checked<% End If %>>
							</td>
						</tr>
						<tr>
						    <td align="right" nowrap><b>Cantidad de Tickets:</b></td>
							<td>
								<input type="text" name="cantic" size="5" maxlength="5" value="<%= l_cantic %>">
							</td>
						    <td align="right" nowrap><b>Cantidad de Remitos:</b></td>
							<td>
								<input type="text" name="canrem" size="5" maxlength="5" value="<%= l_canrem %>">
							</td>
						</tr>
						<tr>
						    <td align="right" nowrap><b>Tipo de Impresión:</b></td>
							<td colspan="3">
								<select name="etiosob" disabled readonly class="deshabinp" size="1" style="width:200;" <%'= l_claseCombo %>>
									<option value="E">Etiqueta</option>
									<option value="S">Sobre</option>
								</select>
								<% 	If l_etiosob = "0" or l_etiosob = "" or IsNull(l_etiosob) then
										l_etiosob = "E"
									end if %>
									<script> document.datos.etiosob.value= "<%= l_etiosob %>";</script>
							</td>
						</tr>
						<tr>
						    <td align="right" nowrap><b>Mostrar Tránsitos:</b></td>
							<td>
								<input type="text" name="mostra" size="12" maxlength="3" value="<%= l_mostra %>">&nbsp;<b>días</b>
							</td>
						</tr>						
						<tr>
						    <td align="right" nowrap><b>Circuito Descarga:</b></td>
							<td>
								<input type="text" name="cirpla" size="12" maxlength="12" value="<%= l_cirpla %>">
							</td>
						</tr>
						
					</table>
				</td>
				<td width="50%"></td>
			</tr>
		</table>
	</td>
</tr>
<iframe name="ifrm2" src="" width="0" height="0" style="visibility:hidden;"></iframe>
</table>
</form>


<%
set l_rs = Nothing
cn.Close
set cn = Nothing
%>
</body>
</html>

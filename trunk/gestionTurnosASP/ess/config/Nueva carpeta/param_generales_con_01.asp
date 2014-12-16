<% Option Explicit %>
<!--#include virtual="/ticket/shared/db/conn_db.inc"-->
<% 
'Archivo: param_generales_con_01.asp
'Descripción: Abm de parametros de generales
'Autor : Lisandro Moro
'Modicado: Raul Chinestra - se agrego el campo de cupo en planta o playa
'Fecha: 28/02/2005

'on error goto 0
Dim l_rs
Dim l_sql
Dim l_filtro
Dim l_orden
Dim l_sqlfiltro
Dim l_sqlorden
Dim l_todos

Dim l_empnro 
Dim l_lugnro 
Dim l_entnro 
Dim l_vennro
Dim l_recnro 
'Dim l_humcam 
'Dim l_humvag 
'Dim l_humdir 
Dim l_dismerhum 
Dim l_traemp 
Dim l_balcodcon 
Dim l_pesdespro 
Dim l_txtpla 
Dim l_txtcup 
Dim l_promov 
Dim l_propla 
Dim l_protra 
Dim l_desnro 
Dim l_mosrec 
Dim l_mosnrotap 
Dim l_mosmez 
Dim l_cupplaya
Dim l_txtcos
Dim l_meralkilo
Dim l_carporfersuc
Dim l_carporfernum
Dim l_movcod
Dim l_nroins

Set l_rs = Server.CreateObject("ADODB.RecordSet")

l_sql = "SELECT empnro ,lugnro ,entnro ,recnro, vennro ,humcam ,humvag ,humdir ,dismerhum ,traemp  "
l_sql  = l_sql  & " ,balcodcon ,pesdespro ,txtpla ,txtcup ,promov ,propla ,protra ,desnro  "
l_sql  = l_sql  & " ,mosrec ,mosnrotap ,mosmez, cupplaya, meralkilo, carporfersuc, carporfernum, txtcos "
l_sql  = l_sql  & " ,nroins "
l_sql  = l_sql  & " FROM tkt_config "
rsOpen l_rs, cn, l_sql, 0
if not l_rs.eof then
	l_empnro = l_rs("empnro")
	l_lugnro = l_rs("lugnro")
	l_entnro = l_rs("entnro")
	l_vennro = l_rs("vennro")
	l_recnro = l_rs("recnro")
'	l_humcam = l_rs("humcam")
'	l_humvag = l_rs("humvag")
'	l_humdir = l_rs("humdir")
	l_dismerhum = l_rs("dismerhum")
	l_traemp = l_rs("traemp")
	l_balcodcon = l_rs("balcodcon")
	l_pesdespro = l_rs("pesdespro")
	l_txtpla = l_rs("txtpla")
	l_txtcup = l_rs("txtcup")
	l_promov = l_rs("promov")
	l_propla = l_rs("propla")
	l_protra = l_rs("protra")
	l_desnro = l_rs("desnro")
	l_mosrec = l_rs("mosrec")
	l_mosnrotap = l_rs("mosnrotap")
	l_mosmez = l_rs("mosmez")
	l_cupplaya = l_rs("cupplaya")
	l_txtcos = l_rs("txtcos")
	l_meralkilo = l_rs("meralkilo")
	l_carporfersuc = l_rs("carporfersuc")
	l_carporfernum = l_rs("carporfernum")
	l_nroins = l_rs("nroins")
else
	l_empnro = 0
	l_lugnro = 0
	l_entnro = 0
	l_vennro = 0
	l_recnro = 0
'	l_humcam = 0
'	l_humvag = 0
'	l_humdir = ""
	l_dismerhum = 0
	l_traemp = 0
	l_balcodcon = 0
	l_pesdespro = ""
	l_txtpla = ""
	l_txtcup = ""
	l_promov = 0
	l_propla = 0
	l_protra = 0
	l_desnro = 0
	l_mosrec = 0
	l_mosnrotap = 0
	l_mosmez = 0
	l_cupplaya = 0
	l_txtcos = ""
	l_meralkilo = 0	
	l_carporfersuc = 0
	l_carporfernum = 0
	l_nroins = 0
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
<title><%= Session("Titulo")%>Parámetros de Generales - Ticket</title>
</head>
<script>
function Valida(){
	document.datos.balcodcon.value = Trim(document.datos.balcodcon.value);
	//document.datos.promov.value = Trim(document.datos.promov.value);
	document.datos.propla.value = Trim(document.datos.propla.value);
	document.datos.protra.value = Trim(document.datos.protra.value);

	/*if(!rutaArchivoValido(document.datos.humdir.value)){	
		document.datos.humdir.select();
		alert('Debe ingresar una ruta válida \n en Directorio Humedímetro.');
		document.datos.humdir.focus();
		return;
	}*/
	if(!rutaArchivoValido(document.datos.txtcup.value)){	
		document.datos.txtcup.select();
		alert('Debe ingresar una ruta válida \n en Archivo para Cupos.');
		document.datos.txtcup.focus();
		return;
	}

	if (isNaN(document.datos.balcodcon.value)){
		document.datos.balcodcon.select();
		alert('Debe ingresar un valor numérico \n en Código Control Balanza.');
		document.datos.balcodcon.focus();
		return;
	}

	if(!rutaArchivoValido(document.datos.txtpla.value)){	
		document.datos.txtpla.select();
		alert('Debe ingresar una ruta válida \n en Archivo para Playa.');
		document.datos.txtpla.focus();
		return;
	}
	if(!rutaArchivoValido(document.datos.txtcos.value)){	
		document.datos.txtcos.select();
		alert('Debe ingresar una ruta válida \n en Directorio para Cospel.');
		document.datos.txtcos.focus();
		return;
	}

	/*if (isNaN(document.datos.promov.value)){
		document.datos.promov.select();
		alert('Debe ingresar un valor numérico \n en Próximo Número de Movimiento.');
		document.datos.promov.focus();
		return;
	}*/
	if (isNaN(document.datos.propla.value)){
		document.datos.propla.select();
		alert('Debe ingresar un valor numérico \n en Proximo Número de Playa.');
		document.datos.propla.focus();
		return;
	}
	if (isNaN(document.datos.protra.value)){
		document.datos.protra.select();
		alert('Debe ingresar un valor numérico \n en Próximo Número de Transile.');
		document.datos.protra.focus();
		return;
	}
	if (document.datos.carporfersuc.value == ''){
		alert('Debe ingresar un valor numérico \n en la Sucursal de la Carta de Porte del Ferrocarril.');
		document.datos.carporfersuc.focus();
		return;
	}
	
	if (document.datos.carporfernum.value == ''){
		alert('Debe ingresar un valor numérico \n en el Número de la Carta de Porte del Ferrocarril.');
		document.datos.carporfernum.focus();
		return;
	}
	
	//abrirVentana('',"newa",500,500);
	//document.datos.target = "newa";
	document.datos.submit();
}
</script>

<body leftmargin="0" rightmargin="0" topmargin="0" bottommargin="0" >
<form name="datos" action="param_generales_con_03.asp" method="post">
<table cellspacing="0" cellpadding="0" border="0" width="100%" height="100%">
<tr>
	<td colspan="2" height="100%">
		<table border="0" cellspacing="0" cellpadding="0">
			<tr>
				<td width="50%"></td>
				<td>
					<table cellspacing="0" cellpadding="0" border="0">
						<tr>
						    <td align="right" nowrap><b>Empresa:</b></td>
							<td colspan="3">
								<select name="empnro" size="1" style="width:100%;" <%'= l_claseCombo %>>
									<option value=0 selected>&laquo; Seleccione una Empresa &raquo;</option>
								<%	l_sql = "SELECT empnro, empdes, empcod "
									l_sql  = l_sql  & " FROM tkt_empresa "
									l_sql  = l_sql  & " ORDER BY empdes "
									rsOpen l_rs, cn, l_sql, 0
									do until l_rs.eof %>	
									<option value=<%= l_rs("empnro") %> > 
									<%= l_rs("empdes") %> (<%=l_rs("empcod")%>) </option>
									<%	l_rs.Movenext
									loop
									l_rs.Close %>
								</select>
								<% If l_empnro = "0" or l_empnro = "" or IsNull(l_empnro) then
										l_empnro = 0
								   end if %>
								   <script> document.datos.empnro.value= "<%= l_empnro %>"</script>
							</td>
						</tr>
						<tr>
							<td align="right" nowrap><b>Lugar:</b></td>
							<td colspan="3">
								<select name="lugnro" size="1" style="width:100%;" <%'= l_claseCombo %>>
									<option value=0 selected>&laquo; Seleccione un Lugar &raquo;</option>
								<%	l_sql = "SELECT lugnro, lugdes, lugcod "
									l_sql  = l_sql  & " FROM tkt_lugar "
									l_sql  = l_sql  & " ORDER BY lugdes "
									rsOpen l_rs, cn, l_sql, 0
									do until l_rs.eof %>	
									<option value=<%= l_rs("lugnro") %> > 
									<%= l_rs("lugdes") %> (<%=l_rs("lugcod")%>) </option>
									<%	l_rs.Movenext
									loop
									l_rs.Close %>
								</select>
								<% If l_lugnro = "0" or l_lugnro = "" or IsNull(l_lugnro) then
									l_lugnro = 0
								end if %>
									<script> document.datos.lugnro.value= "<%= l_lugnro %>"</script>
							</td>
						</tr>
						<tr>
						    <td align="right" nowrap><b>Entregador por Defecto:</b></td>
							<td colspan="3">
								<select name="entnro" size="1" style="width:100%;" <%'= l_claseCombo %>>
									<option value=0 selected>&laquo; Seleccione un Entregador &raquo;</option>
								<%	l_sql = "SELECT entnro, entdes, entcod "
									l_sql  = l_sql  & " FROM tkt_entrec "
									l_sql  = l_sql  & " WHERE (entrol = 'E' OR entrol = 'A') "
									l_sql  = l_sql  & " AND entact = -1 "
									l_sql  = l_sql  & " ORDER BY entdes "
									rsOpen l_rs, cn, l_sql, 0
									do until l_rs.eof %>	
									<option value=<%= l_rs("entnro") %> > 
									<%= l_rs("entdes") %> (<%=l_rs("entcod")%>) </option>
									<%	l_rs.Movenext
									loop
									l_rs.Close %>
								</select>
								<% If l_entnro = "0" or l_entnro = "" or IsNull(l_entnro) then
									l_entnro = 0
								end if %>
									<script> document.datos.entnro.value= "<%= l_entnro %>"</script>
							</td>
						</tr>
						<tr>
							<td align="right" nowrap><b>Recibidor por Defecto:</b></td>
							<td colspan="3">
								<select name="recnro" size="1" style="width:100%;" <%'= l_claseCombo %>>
									<option value=0 selected>&laquo; Seleccione un Entregador &raquo;</option>
								<%	l_sql = "SELECT entnro, entdes, entcod "
									l_sql  = l_sql  & " FROM tkt_entrec "
									l_sql  = l_sql  & " WHERE (entrol = 'R' OR entrol = 'A') "
									l_sql  = l_sql  & " AND entact = -1 "
									l_sql  = l_sql  & " ORDER BY entdes "
									rsOpen l_rs, cn, l_sql, 0
									do until l_rs.eof %>	
									<option value=<%= l_rs("entnro") %> > 
									<%= l_rs("entdes") %> (<%=l_rs("entcod")%>) </option>
									<%	l_rs.Movenext
									loop
									l_rs.Close %>
								</select>
								<% If l_recnro = "0" or l_recnro = "" or IsNull(l_recnro) then
									l_recnro = 0
								end if %>
									<script> document.datos.recnro.value= "<%= l_recnro %>"</script>
							</td>
						</tr>
						<tr>
						    <td align="right" nowrap><b>Destinatario por Defecto:</b></td>
							<td colspan="3">
								<select name="desnro" size="1" style="width:100%;" <%'= l_claseCombo %>>
									<option value=0 selected>&laquo; Seleccione un Destinatario &raquo;</option>
								<%	l_sql = "SELECT vencornro, vencordes "
									l_sql  = l_sql  & " FROM tkt_vencor "
									l_sql  = l_sql  & " WHERE vencortip = 'V' "
									l_sql  = l_sql  & " ORDER BY vencordes "
									rsOpen l_rs, cn, l_sql, 0
									do until l_rs.eof %>	
									<option value=<%= l_rs("vencornro") %> ><%= l_rs("vencordes") %></option>
									<%	l_rs.Movenext
									loop
									l_rs.Close %>
								</select>
								<% If l_desnro = "0" or l_desnro = "" or IsNull(l_desnro) then
									l_desnro = 0
								end if %>
									<script> document.datos.desnro.value= "<%= l_desnro %>"</script>
							</td>
						</tr>
						<tr>
						    <td align="right" nowrap><b>Vendedor por Defecto:</b></td>
							<td colspan="3">
								<select name="vennro" size="1" style="width:100%;" <%'= l_claseCombo %>>
									<option value=0 selected>&laquo; Seleccione un Vendedor &raquo;</option>
								<%	l_sql = "SELECT vencornro, vencordes "
									l_sql  = l_sql  & " FROM tkt_vencor "
									l_sql  = l_sql  & " WHERE vencortip = 'V' "
									l_sql  = l_sql  & " ORDER BY vencordes "
									rsOpen l_rs, cn, l_sql, 0
									do until l_rs.eof %>	
									<option value=<%= l_rs("vencornro") %> ><%= l_rs("vencordes") %></option>
									<%	l_rs.Movenext
									loop
									l_rs.Close %>
								</select>
								<% If l_vennro = "0" or l_vennro = "" or IsNull(l_vennro) then
									l_vennro = 0
								end if %>
									<script> document.datos.vennro.value= "<%= l_vennro %>"</script>
							</td>
						</tr>
						
						<!--tr>
						    <td align="right" nowrap><b>Humedímetro p/Camiónes:</b></td>
							<td>
								<input type="checkbox" name="humcam" <%' If l_humcam = -1 then %>Checked<%' End If %>>
							</td>
						    <td align="right" nowrap><b>Humedímetro p/Vagones:</b></td>
							<td>
								<input type="checkbox" name="humvag"  <%' If l_humvag = -1 then %>Checked<%' End If %>>
							</td>
						</tr>
						<tr>
						    <td align="right" nowrap><b>Directorio Humedímetro:</b></td>
							<td colspan="3">
								<input type="text" name="humdir" size="50" maxlength="50" value="<%'= l_humdir %>">
							</td>
						</tr-->
						<tr>
						    <td align="right" nowrap><b>Discrimina Merma por Humedad:</b></td>
							<td>
								<input type="checkbox" disabled readonly class="deshabinp" name="dismerhum" <% If l_dismerhum = -1 then %>Checked<% End If %>>
							</td>
						    <td align="right" nowrap><b>Pide Empresa en Tránsito:</b></td>
							<td>
								<input type="checkbox" disabled readonly class="deshabinp" name="traemp" <% If l_traemp = -1 then %>Checked<% End If %>>
							</td>
						</tr>
						<tr>
						    <td align="right" nowrap><b>Código Control Balanza:</b></td>
							<td colspan="3">
								<input type="text" disabled readonly class="deshabinp" name="balcodcon" size="4" maxlength="4" value="<%= l_balcodcon %>" >
							</td>
						</tr>
						<tr>
						    <td align="right" nowrap><b>Peso Destino/Procedencia:</b></td>
							<td  colspan="3">
								<select name="pesdespro" disabled readonly class="deshabinp" size="1" style="width:200;" <%'= l_claseCombo %>>
									<option value="0" selected>&laquo; Seleccione un Peso &raquo;</option>
									<option value="D" >Destino (D)</option>
									<option value="P" >Procedencia (P)</option>
								</select>
								<% If l_pesdespro = "0" or l_pesdespro = "" or IsNull(l_pesdespro) then
									l_pesdespro = 0
								end if %>
									<script> document.datos.pesdespro.value= "<%= l_pesdespro %>"</script>
							</td>
						</tr>
						<tr>
						    <td align="right" nowrap><b>Próximo Nro. Carta de Porte Ferro:</b></td>
							<td colspan="3">
								<input type="text" class="habinp" name="carporfersuc" size="4" maxlength="4" value="<%= l_carporfersuc %>">
								<b>-</b>
								<input type="text"  class="habinp" name="carporfernum" size="8" maxlength="8" value="<%= l_carporfernum %>">
							</td>
						</tr>
						<tr>
						    <td align="right" nowrap><b>Archivo para Cupos:</b></td>
							<td colspan="3">
								<input type="text" disabled readonly class="deshabinp" name="txtcup" size="50" maxlength="50" value="<%= l_txtcup %>" >
							</td>
						</tr>
						<tr>
						    <td align="right" nowrap><b>Cupo en Playa:</b></td>
							<td colspan="3">
								<input type="checkbox" name="cuppla" <% If l_cupplaya = -1 then %>Checked<% End If %>>
							</td>
						</tr>

						<tr>
						    <td align="right" nowrap><b>Archivo para Playa:</b></td>
							<td colspan="3">
								<input type="text" name="txtpla" size="50" maxlength="50" value="<%= trim(l_txtpla) %>" >
							</td>
						</tr>
						<tr>
						    <td align="right" nowrap><b>Directorio para Cospel:</b></td>
							<td colspan="3">
								<input type="text" name="txtcos" size="50" maxlength="50" value="<%= trim(l_txtcos) %>" >
							</td>
						</tr>
						
						<tr>
						    <td align="right" nowrap><b>Próximo Número de Movimiento:</b></td>
							<td>
								<%	l_sql = " SELECT openro FROM tkt_operacion "
									rsOpen l_rs, cn, l_sql, 0 
'									response.write(left("000",3- len(l_nroins)) & l_nroins)'  & l_rs("openro"))  & "<br>"
'									response.write(Left("0000000",7 - len(l_rs("openro"))) & l_rs("openro"))
'									response.end
									if not l_rs.eof then
										l_movcod = (left("000",3- len(l_nroins)) & l_nroins) & (Left("0000000",7 - len(l_rs("openro") + 1)) & l_rs("openro") + 1)
									else
										l_movcod =  (left("000",3- len(l_nroins)) & "0000001")
									end if
									%>
								<input type="text" name="movcod" size="10" maxlength="10" value="<%= l_movcod %>" readonly class="deshabinp"> 
							</td>
						    <td align="right" nowrap><b>Próximo Número de Playa:</b></td>
							<td>
								<input type="text" name="propla" size="4" maxlength="4" value="<%= l_propla %>" readonly class="deshabinp">
							</td>
						</tr>
						<tr>
						    <td align="right" nowrap><b>Próximo Número de Transile:</b></td>
							<td colspan="3">
								<input type="text" name="protra" size="4" maxlength="4" value="<%= l_protra %>" readonly class="deshabinp">
							</td>
						</tr>
						<tr>
						    <td align="right" nowrap><b>Muestra el Recibidor:</b></td>
							<td>
								<input type="checkbox" disabled readonly class="deshabinp" name="mosrec" <% If l_mosrec = -1 then %>Checked<% End If %>>
							</td>
						    <td align="right" nowrap><b>Muestra Número de tapas:</b></td>
							<td>
								<input type="checkbox" disabled readonly class="deshabinp" name="mosnrotap" <% If l_mosnrotap = -1 then %>Checked<% End If %>>
							</td>
						</tr>
						<tr>
						    <td align="right" nowrap><b>Muestra la Mezcla:</b></td>
							<td>
								<input type="checkbox"  disabled readonly class="deshabinp" name="mosmez" <% If l_mosmez = -1 then %>Checked<% End If %>>
							<!--/td>
						    <td align="right" nowrap><b>Merma el Kilo:</b></td>
							<td>
								<input type="checkbox"  disabled readonly class="deshabinp" name="meralkilo" <% If l_meralkilo = -1 then %>Checked<% End If %>>
							</td-->
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

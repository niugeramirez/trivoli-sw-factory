<% Option Explicit %>
<!--#include virtual="/ticket/shared/db/conn_db.inc"-->
<% 
'Archivo: param_calidad_con_01.asp
'Descripción: Abm de parametros de calidad
'Autor : Lisandro Moro
'Fecha: 28/02/2005

on error goto 0
Dim l_rs
Dim l_sql
Dim l_filtro
Dim l_orden
Dim l_sqlfiltro
Dim l_sqlorden
Dim l_todos

Dim l_camnro
Dim l_muecam
Dim l_etidescamsec
Dim l_etidescamhum
Dim l_etisal
Dim l_eticie
Dim l_paleo
Dim l_fumiga
Dim l_secado
Dim l_meralkilo
Dim l_tramuecam
Dim l_traimpeti

l_filtro = request("filtro")
l_orden  = request("orden")

Set l_rs = Server.CreateObject("ADODB.RecordSet")

l_sql = "SELECT camnro, muecam, etidescamsec, etidescamhum, etisal, eticie,  "
l_sql  = l_sql  & " paleo, fumiga, secado, meralkilo, tramuecam, traimpeti "
l_sql  = l_sql  & " FROM tkt_config "
rsOpen l_rs, cn, l_sql, 0
if not l_rs.eof then
	l_camnro = l_rs("camnro")
	l_muecam = l_rs("muecam")
	l_etidescamsec = l_rs("etidescamsec")
	l_etidescamhum = l_rs("etidescamhum")
	l_etisal = l_rs("etisal")
	l_eticie = l_rs("eticie")
	l_paleo = l_rs("paleo")
	l_fumiga = l_rs("fumiga")
	l_secado = l_rs("secado")
	l_meralkilo = l_rs("meralkilo")
	l_tramuecam = l_rs("tramuecam")
	l_traimpeti = l_rs("traimpeti")
else
	l_camnro = 0
	l_muecam = 0
	l_etidescamsec = 0
	l_etidescamhum = 0
	l_etisal = 0
	l_eticie = 0
	l_paleo = 0
	l_fumiga = 0
	l_secado = 0
	l_meralkilo = 0
	l_tramuecam = 0
	l_traimpeti = 0
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
<title><%= Session("Titulo")%>Parámetros de Calidad - Ticket</title>
</head>

<script>
function Valida(){
	document.datos.etidescamsec.value = Trim(document.datos.etidescamsec.value);
	document.datos.etidescamhum.value = Trim(document.datos.etidescamhum.value);
	document.datos.etisal.value = Trim(document.datos.etisal.value);
	document.datos.eticie.value = Trim(document.datos.eticie.value);
	
	if (isNaN(document.datos.etidescamsec.value)){
		document.datos.etidescamsec.select();
		alert('Debe ingresar un valor numérico en \n Cantidad Etiquetas p/Seco');
		document.datos.etidescamsec.focus();
		return;
	}
	if (isNaN(document.datos.etidescamhum.value)){
		document.datos.etidescamhum.select();
		alert('Debe ingresar un valor numérico \n en Cantidad Etiquetas p/Humedo');
		document.datos.etidescamhum.focus();
		return;
	}
	if (isNaN(document.datos.etisal.value)){
		document.datos.etisal.select();
		alert('Debe ingresar un valor numérico \n en Cantidad Etiquetas Salida');
		document.datos.etisal.focus();
		return;
	}
	if (isNaN(document.datos.eticie.value)){
		document.datos.eticie.select();
		alert('Debe ingresar un valor numérico \n en Cantidad Etiquetas Cierre');
		document.datos.eticie.focus();
		return;
	}
	//abrirVentana('',"newa",500,500);
	//document.datos.target = "newa";
	document.datos.submit();
}
</script>

<body leftmargin="0" rightmargin="0" topmargin="0" bottommargin="0">
<form name="datos" action="param_calidad_con_03.asp" method="post" target="ifrm2">
<table cellspacing="0" cellpadding="0" border="0" width="100%" height="100%">
<tr>
	<td colspan="2" height="100%">
		<table border="0" cellspacing="0" cellpadding="0">
			<tr>
				<td width="50%"></td>
				<td>
					<table cellspacing="0" cellpadding="0" border="0">
						<tr>
						    <td align="right" nowrap><b>Cámara Arbitral:</b></td>
							<td colspan="3">
								<select name="camnro" size="1" style="width:200;" <%'= l_claseCombo %>>
									<option value=0 selected>&laquo; Seleccione una Cámara &raquo;</option>
								<%	l_sql = "SELECT camnro, camdes, camcod "
									l_sql  = l_sql  & " FROM tkt_camara "
									l_sql  = l_sql  & " ORDER BY camdes "
									rsOpen l_rs, cn, l_sql, 0
									do until l_rs.eof %>	
									<option value=<%= l_rs("camnro") %> > 
									<%= l_rs("camdes") %> (<%=l_rs("camcod")%>) </option>
									<%	l_rs.Movenext
									loop
									l_rs.Close %>
								</select>
								<% If l_camnro = "0" or l_camnro = "" or IsNull(l_camnro) then
									l_camnro = 0
								end if %>
									<script> document.datos.camnro.value= "<%= l_camnro %>"</script>
							</td>
						</tr>
						<tr>
						    <td align="right" nowrap><b>Muestra a Cámara:</b></td>
							<td colspan="3">
								<input type="checkbox" name="muecam" <% If l_muecam = -1 then %>Checked<% End If %>>
							</td>
						</tr>
						<tr>
						    <td align="right" nowrap><b>Cantidad Etiquetas p/Seco:</b></td>
							<td>
								<input type="text" name="etidescamsec" size="2" maxlength="2" value="<%= l_etidescamsec %>" >
							</td>
						    <td align="right" nowrap><b>Paleo:</b></td>
							<td>
								<input type="checkbox" name="paleo" <% If l_paleo = -1 then %>Checked<% End If %>>
							</td>
						</tr>
						<tr>
						    <td align="right" nowrap><b>Cantidad Etiquetas p/Humedo:</b></td>
							<td>
								<input type="text" name="etidescamhum" size="2" maxlength="2" value="<%= l_etidescamhum %>" >
							</td>
						    <td align="right" nowrap><b>Fumigación:</b></td>
							<td>
								<input type="checkbox" name="fumiga" <% If l_fumiga = -1 then %>Checked<% End If %>>
							</td>
						</tr>
						<tr>
						    <td align="right" nowrap><b>Cantidad Etiquetas Salida:</b></td>
							<td>
								<input type="text" name="etisal" size="2" maxlength="2" value="<%= l_etisal %>" >
							</td>
						    <td align="right" nowrap><b>Secado:</b></td>
							<td>
								<input type="checkbox" name="secado" <% If l_secado = -1 then %>Checked<% End If %>>
							</td>
						</tr>
						<tr>
						    <td align="right" nowrap><b>Cantidad Etiquetas Cierre:</b></td>
							<td>
								<input type="text" name="eticie" size="2" maxlength="2" value="<%= l_eticie %>" >
							</td>
						    <td align="right" nowrap><b>Mermas al Kilo:</b></td>
							<td>
								<input type="checkbox" name="meralkilo" <% If l_meralkilo = -1 then %>Checked<% End If %>>
							</td>
						</tr>


						<tr>
						    <td align="right" nowrap><b>Tránsito envía Muesta a Cámara:</b></td>
							<td>
								<input type="checkbox" disabled readonly class="deshabinp" name="tramuecam" <% If l_tramuecam = -1 then %>Checked<% End If %>>
							</td>
						    <td align="right" lass="deshabinp" nowrap><b>Tránsito Imprime Etiquetas:</b></td>
							<td>
								<input type="checkbox" disabled readonly class="deshabinp" name="traimpeti" <% If l_traimpeti = -1 then %>Checked<% End If %>>
							</td>
						</tr>

					</table>
				</td>
				<td width="50%"></td>
			</tr>
		</table>
	</td>
</tr>

</table>
<iframe name="ifrm2" src="" width="0" height="0" style="visibility:hidden;"></iframe>
</form>


<%
set l_rs = Nothing
cn.Close
set cn = Nothing
%>
</body>
</html>

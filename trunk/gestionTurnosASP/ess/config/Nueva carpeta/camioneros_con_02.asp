<% Option Explicit %>
<!--#include virtual="/serviciolocal/shared/db/conn_db.inc"-->
<% 
'Archivo:camioneros_con_02.asp
'Descripción: Abm de camioneros
'Autor : Lisandro Moro
'Fecha: 15/02/2005
'Modificado: Raul Chinestra - 16/03/2005 - Se controlo que el codigo sea numerico

'Datos del formulario
'
on error goto 0
Dim l_camcod
Dim l_camdes
Dim l_camcha
Dim l_camaco
Dim l_camcas
Dim l_tranro
Dim l_camact
Dim l_camhab
Dim l_camsis

								
Dim l_tipo
Dim l_trades
Dim l_camcuit
Dim l_camnro
Dim l_clase
Dim l_claseChek
Dim l_claseCombo
'ADO
Dim l_sql
Dim l_rs

l_tipo = request.querystring("tipo")
l_camnro = request.querystring("cabnro")
%>
<html>
<head>
<link href="/serviciolocal/shared/css/tables3.css" rel="StyleSheet" type="text/css">
<meta http-equiv="Content-Type" content="text/html; charset=iso-8859-1">
<title><%= Session("Titulo")%>Camioneros - Ticket</title>
</head>
<style type="text/css">
.none{
	padding : 0;
	padding-left : 0;
}
</style>
<script src="/serviciolocal/shared/js/fn_ayuda.js"></script>
<script src="/serviciolocal/shared/js/fn_windows.js"></script>
<script src="/serviciolocal/shared/js/fn_valida.js"></script>
<script>

function Valida(){
	<% If l_tipo = "A" OR l_tipo = "M"then%>
	if (document.datos.camcod.value == ""){
		alert('Debe ingresar un Código.');
		document.datos.camcod.focus();
		return;
	}
	/*if (document.datos.nrodoc.value == ""){
		alert('Debe ingresar un Cuit.');
		document.datos.nrodoc.focus();
		return;
	}
	if (!ValidaCuit(document.datos.nrodoc.value)){
		document.datos.nrodoc.focus();
		return;
	}*/	
	if (isNaN(document.datos.camcod.value)){
		alert("El Código ingresado debe ser Numérico");
		document.datos.camcod.focus();
		return;
	}
	if(document.datos.camdes.value == ""){
		alert("Debe ingresar una descripción");
		document.datos.camdes.focus();
		return;
	}
	if(document.datos.camcha.value == ""){
		alert("Debe ingresar la Patente del Chasis. ");
		document.datos.camcha.focus();
		return;
	}
	if (document.datos.camcas.checked){
		document.datos.camcas.value = -1;
	}else{
		document.datos.camcas.value = 0;
	}
	if (document.datos.camhab.checked){
		document.datos.camhab.value = -1;
	}else{
		document.datos.camhab.value = 0;
	}

	document.datos.submit();
		<% If l_tipo = "M"  then %>
		//	abrirVentanaH('camioneros_con_03.asp?tipo=<%= l_tipo %>&cabnro=' + document.datos.camnro.value,'',520,160);
		<% Else  %>	
		//	abrirVentanaH('camioneros_con_03.asp?tipo=<%= l_tipo %>','',520,160);
		<% End If %>
	<% Else  %>
	window.close();
	<% End If %>
}
</script>
<% 
Select case l_tipo
	case "M" , "C"
		Set l_rs = Server.CreateObject("ADODB.RecordSet")
		l_sql = "SELECT camcod, camdes, camcha, camaco, camcas, camact, camhab, camsis "'nrodoc, , tkt_camionero.tranro,  trades
		l_sql = l_sql & " FROM tkt_camionero "
		l_sql = l_sql & " LEFT JOIN tkt_terdoc ON tkt_terdoc.valnro = tkt_camionero.camnro AND tipternro = 2 AND tipdocnro = 5 "
		'l_sql = l_sql & " LEFT JOIN tkt_transportista ON tkt_transportista.tranro = tkt_camionero.tranro"
		l_sql = l_sql & " WHERE camnro = " & l_camnro
		rsOpen l_rs, cn, l_sql, 0 
		if not l_rs.eof then
			l_camcod = l_rs("camcod")
			l_camdes = l_rs("camdes")
			l_camcha = l_rs("camcha")
			l_camaco = l_rs("camaco")
			l_camcas = l_rs("camcas")
'			l_tranro = l_rs("tranro")
			l_camact = l_rs("camact")
			l_camhab = l_rs("camhab")
			l_camsis = l_rs("camsis")
'			l_trades = l_rs("trades")
			'l_camcuit =  l_rs("nrodoc")
			if l_camsis = -1 then
				l_tipo = "C"
			end if
		end if
		l_rs.Close
	case "A"
		l_camcod = ""
		l_camdes = ""
		l_camcha = ""
		l_camaco = ""
		l_camcas = ""
'		l_tranro = 0
		l_camact = ""
		l_camhab = ""
		l_camsis = ""
'		l_trades = ""
		l_camcuit =  ""
end Select

Select case l_tipo
	case "A"
		l_clase = " class=""habinp"" "
		l_claseChek = " "
		l_claseCombo = " class=""habinp"" "
	case "M"
		l_clase = " class=""habinp"" "
		l_claseChek = " "
		l_claseCombo = " class=""habinp"" "
	case "C"
		l_clase = " class=""deshabinp"" readonly "
		l_claseChek = " readonly disabled"
		l_claseCombo = " class=""deshabinp"" disabled"
end Select
%>
<body leftmargin="0" rightmargin="0" topmargin="0" bottommargin="0" onload="document.datos.camcod.focus();">
<form name="datos" action="camioneros_con_03.asp?tipo=<%= l_tipo %>&cabnro=<%= l_camnro %>" method="post" target="ifrm">

<table cellspacing="0" cellpadding="0" border="0" width="100%" height="100%">
<tr>
    <td class="th2"  nowrap>Camioneros</td>
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
						    <td align="right" nowrap><b>Código:</b></td>
							<td>
								<input type="Text" name="camcod" size="8" maxlength="8" value="<%= l_camcod %>" <%= l_clase %>>
							</td>
							<td align="right" nowrap><!--<b>Cuit:</b>--></td>
							<td>
								<!--<input type="Text" name="nrodoc" size="30" maxlength="30" value="<%'= l_camcuit %>" <%= l_clase %>>-->
							</td>
						</tr>
						<tr>
							<td align="right" nowrap><b>Apellido y Nombre:</b></td>
							<td colspan="3">
								<input type="Text" name="camdes" size="50" maxlength="50" value="<%= l_camdes %>" <%= l_clase %>>
							</td>
						</tr>
	<!--					<tr>
						    <td align="right" nowrap><b>Transportista:</b></td>
							<td colspan="3">
								<select name="tranro" size="1" style="width:325;" <%'= l_claseCombo %>>
									<option value=0 selected>&laquo; Seleccione un Transportista &raquo;</option>
									<%'Set l_rs = Server.CreateObject("ADODB.RecordSet")
									'l_sql = "SELECT tranro, trades, tracod "
									'l_sql  = l_sql  & " FROM tkt_transportista "
									'l_sql  = l_sql  & " ORDER BY trades "
									'rsOpen l_rs, cn, l_sql, 0
									'do until l_rs.eof		%>	
									<option value=<%'= l_rs("tranro") %> > 
									<%'= l_rs("trades") %> (<%'=l_rs("tracod")%>) </option>
									<%'	l_rs.Movenext
									'loop
									'l_rs.Close %>
								</select>
								<%' If l_tranro = "0" or l_tranro = "" or IsNull(l_tranro) then
								'	l_tranro = 0
								'end if %>
									<script> document.datos.tranro.value= "<%'= l_tranro %>"</script>
							</td>
						</tr>-->
						<tr>
						    <td align="right" nowrap><b>Patente Chasis:</b></td>
							<td colspan="3">
								<input type="Text" name="camcha" size="8" maxlength="8" value="<%= l_camcha %>" <%= l_clase %>>
							</td>
						</tr>
						<tr>
							<td align="right" nowrap><b>Patente Acoplado:</b></td>
							<td colspan="3">
								<input type="Text" name="camaco" size="8" maxlength="8" value="<%= l_camaco %>" <%= l_clase %>>
							</td>
						</tr>
						<tr>
						    <td align="right" nowrap><b>Cáscara:</b></td>
							<td colspan="3">
								<input type="Checkbox" name="camcas" value="<%= l_camcas %>" <% If l_camcas = -1 then %>Checked<% End If %> <%= l_claseChek %>>
							</td>
						</tr>
						<tr>
							<td align="right" nowrap><b>Habilitado:</b></td>
							<td colspan="3">
								<input type="Checkbox" name="camhab" value="<%= l_camhab %>" <% If l_camhab = -1 or l_tipo = "A" then %>Checked<% End If %> <%= l_claseChek %>>
							</td>
						</tr>
						<tr>
							<td colspan="4" class="barra">Transportistas</td>
						</tr>
						<tr  class="none">
							<td colspan="4" class="none">
								<iframe width="100%" height="100%" name="ifrm2" src="camionero_transportistas_con_01.asp?camnro=<%= l_camnro %>"></iframe>
							</td>
						</tr>
					</table>
				</td>
				<td width="50%"></td>
			</tr>
		</table>
	</td>
</tr>
</form>
<tr>
    <td colspan="2" align="right" class="th2">
		<iframe name="ifrm" width="0" height="0" style="visibility:hidden;" ></iframe>
		<% If l_tipo <> "C" then %>
			<% call MostrarBoton ("sidebtnABM", "Javascript:Valida();","Aceptar")%>
			<a class=sidebtnABM href="Javascript:window.close();">Cancelar</a>
		<% Else  %>
			<a class=sidebtnABM href="Javascript:window.close();">Aceptar</a>
		<% End If %>
	</td>
</tr>
</table>

<%
set l_rs = nothing
Cn.Close
'Cn = nothing
%>
</body>
</html>
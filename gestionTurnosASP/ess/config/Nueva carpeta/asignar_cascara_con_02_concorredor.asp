<% Option Explicit %>
<!--#include virtual="/serviciolocal/shared/db/conn_db.inc"-->
<% 

on error goto 0


'Archivo: asignar_cascara_con_02.asp
'Descripci�n: ABM de Asignaci�n de Nros de C�scara
'Autor : Raul Chinestra
'Fecha: 09/05/3005

'Datos del formulario
Dim l_asicasnro
Dim l_ordnro
Dim l_tarnro
Dim l_camnro
Dim l_tranro
Dim l_camcha
Dim l_camaco
Dim l_vennro
Dim l_entnro
Dim l_cornro
Dim l_desnro
Dim l_deporinro
Dim l_depdesnro


'ADO
Dim l_tipo
Dim l_sql
Dim l_rs

l_tipo = request.querystring("tipo")

%>
<html>
<head>
<link href="/serviciolocal/shared/css/tables3.css" rel="StyleSheet" type="text/css">
<meta http-equiv="Content-Type" content="text/html; charset=iso-8859-1">
<title><%= Session("Titulo")%>Asignaci�n de Nros de C�scara - Ticket</title>
</head>
<script src="/serviciolocal/shared/js/fn_ayuda.js"></script>
<script src="/serviciolocal/shared/js/fn_windows.js"></script>
<script src="/serviciolocal/shared/js/fn_valida.js"></script>
<script>
function Validar_Formulario(){
if (document.datos.ordnro.value == 0){
	alert("Debe ingresar una Orden de Trabajo.");
	document.datos.ordnro.focus();
	return;
}
if (document.datos.tarnro.value == ""){
	alert("Debe ingresar el Nro. de Tarjeta.");
	document.datos.tarnro.focus();
	return;
}
if (document.camionero.datos.camnro.value == 0){
	alert("Debe ingresar el Camionero.");
	document.camionero.datos.camnro.focus();
	return;
}

if (document.datos.camcha.value == ""){
	alert("Debe ingresar la Patente del Chasis.");
	document.datos.camcha.focus();
	return;
}
if (document.datos.tranro.value == 0){
	alert("Debe ingresar la Empresa Transportista.");
	document.datos.tranro.focus();
	return;
}

if (document.datos.vennro.value == 0){
	alert("Debe ingresar Vendedor.");
	document.datos.vennro.focus();
	return;
}

if (document.datos.cornro.value == 0){
	alert("Debe ingresar el Corredor.");
	document.datos.cornro.focus();
	return;
}

if (document.datos.entnro.value == 0){
	alert("Debe ingresar el Entregador.");
	document.datos.entnro.focus();
	return;
}

if (document.datos.desnro.value == 0){
	alert("Debe ingresar el Destinatario.");
	document.datos.desnro.focus();
	return;
}

if (document.datos.deporinro.value == 0){
	alert("Debe ingresar el Dep�sito Origen.");
	document.datos.deporinro.focus();
	return;
}

if (document.datos.depdesnro.value == 0){
	alert("Debe ingresar el Dep�sito Destino.");
	document.datos.depdesnro.focus();
	return;
}


	document.valida.location = "asignar_cascara_con_06.asp?tipo=<%= l_tipo%>&asicasnro="+document.datos.asicasnro.value + "&ordnro="+document.datos.ordnro.value + "&tarnro="+document.datos.tarnro.value  + "&camnro="+document.camionero.datos.camnro.value;

}

function valido(){
	document.datos.submit();
}

function invalido(texto){
	alert(texto);
	document.datos.tarnro.focus();	
}

function CargarCamioneros(){
	document.camionero.location = "asignar_cascara_con_08.asp?ordnro="+ document.datos.ordnro.value;	
	document.datos.camcha.value = "";
	document.datos.camaco.value = "";	
	document.datos.tranro.value = 0;
}

</script>
<% 

Set l_rs = Server.CreateObject("ADODB.RecordSet")

select Case l_tipo
	Case "A":
		l_ordnro  = 0
		l_tarnro = ""
		l_camcha = ""
		l_camaco = ""
		l_tarnro = ""
		l_camnro = 0
		l_tranro = 0
		l_vennro = 0
		l_entnro = 0
		l_cornro = 0
		l_desnro = 0
		l_deporinro = 0
		l_depdesnro = 0
		
	Case "M":
		Set l_rs = Server.CreateObject("ADODB.RecordSet")
		l_asicasnro = request.querystring("cabnro")
		l_sql = "SELECT * "
		l_sql = l_sql & " FROM tkt_asicas "
		l_sql  = l_sql  & " WHERE asicasnro = " & l_asicasnro
		rsOpen l_rs, cn, l_sql, 0 
		if not l_rs.eof then
			l_ordnro = l_rs("ordnro")
			l_tarnro = l_rs("tarnro")
			l_camnro = l_rs("camnro")
			l_camcha = l_rs("camcha")
			l_camaco = l_rs("camaco")
			l_tranro = l_rs("tranro")
			l_vennro = l_rs("vennro")
			l_entnro = l_rs("entnro")
			l_cornro = l_rs("cornro")
			l_desnro = l_rs("desnro")
			l_deporinro = l_rs("deporinro")
			l_depdesnro = l_rs("depdesnro")
		end if
		l_rs.Close
end select
%>
<body leftmargin="0" rightmargin="0" topmargin="0" bottommargin="0" onload="JavaScript:document.datos.ordnro.focus()">
<form name="datos" action="asignar_cascara_con_03.asp?tipo=<%= l_tipo %>" method="post" target="valida1">
<input type="hidden" name="asicasnro" value="<%= l_asicasnro %>">
<input type="hidden" name="camnro" value="<%= l_camnro %>">


<table cellspacing="0" cellpadding="0" border="0" width="100%" height="100%">
<tr>
    <td class="th2" nowrap>&nbsp;</td>
	<td class="th2" align="right">
		<a class=sidebtnHLP href="Javascript:ayuda('<%= Request.ServerVariables("SCRIPT_NAME")%>');">Ayuda</a>
	</td>
</tr>
<tr>
	<td colspan="2" height="100%">
		<table border="0" cellspacing="0" cellpadding="0">
			<tr>
				<td width="20%"></td>
				<td width="60%">
					<table cellspacing="0" cellpadding="0" border="0">
					<tr>
						    <td align="right" nowrap><b>Orden Trabajo:</b></td>
							<td colspan="3">
								<select name="ordnro" size="1" style="width:300;" onchange="Javascript:CargarCamioneros();">
									<option value=0 selected>&laquo; Seleccione una Orden de Trabajo &raquo;</option>
								<%	l_sql = "SELECT ordnro, ordcod "
									l_sql  = l_sql  & " FROM tkt_ordentrabajo "
									'l_sql  = l_sql  & " WHERE (pronro = 1564) "
									l_sql  = l_sql  & " ORDER BY ordcod "
									rsOpen l_rs, cn, l_sql, 0
									do until l_rs.eof %>	
									<option value=<%= l_rs("ordnro") %> > 
									<%= l_rs("ordcod") %> </option>
									<%	l_rs.Movenext
									loop
									l_rs.Close %>
								</select>
								<script> document.datos.ordnro.value= "<%= l_ordnro %>"</script>
							</td>
					</tr>
					<tr>
					    <td align="right"><b>Tarjeta Nro:</b></td>
						<td>
							<input type="text" name="tarnro" size="8" maxlength="5" value="<%= l_tarnro %>">
						</td>
					</tr>

					<tr>
					    <td align="right" nowrap><b>Camionero:</b></td>
					    <td align="left" >
						<iframe name="camionero" frameborder="0" width="100%" height="23" scrolling="No" src="asignar_cascara_con_08.asp?ordnro=<%= l_ordnro %>&camnro=<%= l_camnro %>"></iframe>						
						</td>
					</tr>
					<tr>
					    <td align="right"><b>Chasis:</b></td>
						<td>
							<input type="text" name="camcha" size="15" maxlength="10" value="<%= l_camcha %>">
						</td>
					</tr>																		
					<tr>
					    <td align="right"><b>Acoplado:</b></td>
						<td>
							<input type="text" name="camaco" size="15" maxlength="10" value="<%= l_camaco %>">
						</td>
					</tr>																							
					<tr>
						    <td align="right" nowrap><b>Transportista:</b></td>
							<td colspan="3">
								<select name="tranro" size="1" style="width:300;" >
									<option value=0 selected>&laquo; Seleccione un Transportista &raquo;</option>
								<%	l_sql = "SELECT tranro, trades, tracod "
									l_sql  = l_sql  & " FROM tkt_transportista "
									rsOpen l_rs, cn, l_sql, 0
									do until l_rs.eof %>	
									<option value=<%= l_rs("tranro") %> > 
									<%= l_rs("trades") %> (<%=l_rs("tracod")%>) </option>
									<%	l_rs.Movenext
									loop
									l_rs.Close %>
								</select>
								<script> document.datos.tranro.value="<%= l_tranro %>"</script>
							</td>
					</tr>										
					<tr>
						    <td align="right" nowrap><b>Vendedor:</b></td>
							<td colspan="3">
								<select name="vennro" size="1" style="width:300;">
									<option value=0 selected>&laquo; Seleccione un Vendedor &raquo;</option>
								<%	l_sql = "SELECT vencornro, vencordes, vencorcod "
									l_sql  = l_sql  & " FROM tkt_vencor "
									l_sql  = l_sql  & " WHERE (vencortip = 'V') "
									l_sql  = l_sql  & " AND venact = -1 "
									'l_sql  = l_sql  & " ORDER BY vencordes "
									rsOpen l_rs, cn, l_sql, 0
									do until l_rs.eof %>	
									<option value=<%= l_rs("vencornro") %> > 
									<%= l_rs("vencordes") %> (<%=l_rs("vencorcod")%>) </option>
									<%	l_rs.Movenext
									loop
									l_rs.Close %>
								</select>
								<% If l_vennro = "0" or l_vennro = "" or IsNull(l_vennro) then
									l_vennro = 0
								end if %>								
								<script> document.datos.vennro.value="<%= l_vennro %>"</script>
							</td>
					</tr>				
					<tr>
						    <td align="right" nowrap><b>Corredor:</b></td>
							<td colspan="3">
								<select name="cornro" size="1" style="width:300;">
									<option value=0 selected>&laquo; Seleccione un Corredor &raquo;</option>
								<%	l_sql = "SELECT vencornro, vencordes, vencorcod "
									l_sql  = l_sql  & " FROM tkt_vencor "
									l_sql  = l_sql  & " WHERE (vencortip = 'C') "
									l_sql  = l_sql  & " AND venact = -1 "
									l_sql  = l_sql  & " ORDER BY vencordes "
									rsOpen l_rs, cn, l_sql, 0
									do until l_rs.eof %>	
									<option value=<%= l_rs("vencornro") %> > 
									<%= l_rs("vencordes") %> (<%=l_rs("vencorcod")%>) </option>
									<%	l_rs.Movenext
									loop
									l_rs.Close %>
								</select>
								<% If l_cornro = "0" or l_cornro = "" or IsNull(l_cornro) then
									l_cornro = 0
								end if %>								
								<script> document.datos.cornro.value= "<%= l_cornro %>"</script>
							</td>
					</tr>					
					<tr>
						    <td align="right" nowrap><b>Entregador:</b></td>
							<td colspan="3">
								<select name="entnro" size="1" style="width:300;">
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
						    <td align="right" nowrap><b>Destinatario:</b></td>
							<td colspan="3">
								<select name="desnro" size="1" style="width:300;">
									<option value=0 selected>&laquo; Seleccione un Destinatario &raquo;</option>
								<%	l_sql = "SELECT desnro, desdes "
									l_sql  = l_sql  & " FROM tkt_destinatario "
									l_sql  = l_sql  & " ORDER BY desdes "
									rsOpen l_rs, cn, l_sql, 0
									do until l_rs.eof %>	
									<option value=<%= l_rs("desnro") %> ><%= l_rs("desdes") %></option>
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
						    <td align="right" nowrap><b>Dep�sito Origen:</b></td>
							<td colspan="3">
								<select name="deporinro" size="1" style="width:300;" >
									<option value=0 selected>&laquo; Seleccione un Dep�sito &raquo;</option>
								<%	l_sql = "SELECT depnro, depdes, depcod "
									l_sql  = l_sql  & " FROM tkt_deposito "
									rsOpen l_rs, cn, l_sql, 0
									do until l_rs.eof %>	
									<option value=<%= l_rs("depnro") %> > 
									<%= l_rs("depdes") %> (<%=l_rs("depcod")%>) </option>
									<%	l_rs.Movenext
									loop
									l_rs.Close %>
								</select>
								<script> document.datos.deporinro.value="<%= l_deporinro %>"</script>
							</td>
					</tr>															
					<tr>
						    <td align="right" nowrap><b>Dep�sito Destino:</b></td>
							<td colspan="3">
								<select name="depdesnro" size="1" style="width:300;" >
									<option value=0 selected>&laquo; Seleccione un Dep�sito &raquo;</option>
								<%	l_sql = "SELECT depnro, depdes, depcod "
									l_sql  = l_sql  & " FROM tkt_deposito "
									rsOpen l_rs, cn, l_sql, 0
									do until l_rs.eof %>	
									<option value=<%= l_rs("depnro") %> > 
									<%= l_rs("depdes") %> (<%=l_rs("depcod")%>) </option>
									<%	l_rs.Movenext
									loop
									l_rs.Close %>
								</select>
								<script> document.datos.depdesnro.value="<%= l_depdesnro %>"</script>
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
		<a class=sidebtnABM href="Javascript:Validar_Formulario()">Aceptar</a>
		<a class=sidebtnABM href="Javascript:window.close()">Cancelar</a>
	</td>
</tr>
</table><!-- style="visibility=hidden;" -->
<iframe name="valida"  src="" width="100%" height="100%"></iframe> 
<iframe name="valida1"  src="" width="100%" height="100%"></iframe> 
</form>
<%
set l_rs = nothing
cn.Close
set cn = nothing
%>
</body>
</html>

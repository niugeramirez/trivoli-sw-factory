<% Option Explicit %>
<!--#include virtual="/serviciolocal/shared/db/conn_db.inc"-->
<% 

'on error goto 0


'Archivo: asignar_embarque_con_02.asp
'Descripción: ABM de Asignación de Nros de camioneros de embarque
'Autor : Gustavo Manfrin
'Fecha: 20/09/2006

'Datos del formulario
Dim l_asiembnro
Dim l_tarcod
Dim l_camnro
Dim l_tranro
Dim l_camcha
Dim l_camaco
Dim l_embnro
Dim l_asiembobs

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
<title><%= Session("Titulo")%>Asignación de Nros de Embarque - Ticket</title>
</head>
<script src="/serviciolocal/shared/js/fn_ayuda.js"></script>
<script src="/serviciolocal/shared/js/fn_windows.js"></script>
<script src="/serviciolocal/shared/js/fn_valida.js"></script>
<script>
function Validar_Formulario(){
if (document.datos.tarcod.value == ""){
	alert("Debe ingresar el Nro. de Tarjeta.");
	document.datos.tarcod.focus();
	return;
}
//if ((document.datos.camcha.value.length < 6) || 
//    (document.datos.camcha.value.length > 7) || 
//	(!(isNaN(Left(document.datos.camcha.value,3)))) || 
//	(isNaN(Right(document.datos.camcha.value,3)))){
//	alert("Error la Patente del Chasis.");
//	document.datos.camcha.focus();
//	return;
//}

if ((document.datos.camcha.value.length < 6) ||
    (document.datos.camcha.value.length > 7)) {
	alert("Error la Patente del Chasis.");
	document.datos.camcha.focus();
	return;
}

if ((document.datos.camcha.value.length = 6) && 
	((!(isNaN(document.datos.camcha.value.substr(0,1)))) ||
	(!(isNaN(document.datos.camcha.value.substr(1,1)))) || 
	(!(isNaN(document.datos.camcha.value.substr(2,1)))) || 
    (isNaN(document.datos.camcha.value.substr(3,1))) || 	
    (isNaN(document.datos.camcha.value.substr(4,1))) || 	
    (isNaN(document.datos.camcha.value.substr(5,1))))) {  	
	alert("Error la Patente del Chasis..");
	document.datos.camcha.focus();
	return;
}

if ((document.datos.camcha.value.length = 7) && 
	((!(isNaN(document.datos.camcha.value.substr(0,1)))) ||
	(!(isNaN(document.datos.camcha.value.substr(1,1)))) || 
	(!(isNaN(document.datos.camcha.value.substr(2,1)))) || 
    (isNaN(document.datos.camcha.value.substr(4,1))) || 	
    (isNaN(document.datos.camcha.value.substr(5,1))) || 	
    (isNaN(document.datos.camcha.value.substr(6,1))))) {  	
	alert("Error la Patente del Chasis...");
	document.datos.camcha.focus();
	return;
}

if ((document.datos.camaco.value.length < 6) ||
    (document.datos.camaco.value.length > 7)) {
	alert("Error la Patente del Acoplado.");
	document.datos.camaco.focus();
	return;
}

if ((document.datos.camaco.value.length = 6) && 
	((!(isNaN(document.datos.camaco.value.substr(0,1)))) ||
	(!(isNaN(document.datos.camaco.value.substr(1,1)))) || 
	(!(isNaN(document.datos.camaco.value.substr(2,1)))) || 
    (isNaN(document.datos.camaco.value.substr(3,1))) || 	
    (isNaN(document.datos.camaco.value.substr(4,1))) || 	
    (isNaN(document.datos.camaco.value.substr(5,1))))) {  	
	alert("Error la Patente del Acoplado..");
	document.datos.camaco.focus();
	return;
}

if ((document.datos.camaco.value.length = 7) && 
	((!(isNaN(document.datos.camaco.value.substr(0,1)))) ||
	(!(isNaN(document.datos.camaco.value.substr(1,1)))) || 
	(!(isNaN(document.datos.camaco.value.substr(2,1)))) || 
    (isNaN(document.datos.camaco.value.substr(4,1))) || 	
    (isNaN(document.datos.camaco.value.substr(5,1))) || 	
    (isNaN(document.datos.camaco.value.substr(6,1))))) {  	
	alert("Error la Patente del Acoplado...");
	document.datos.camaco.focus();
	return;
}

if (document.datos.tranro.value == 0){
	alert("Debe ingresar la Empresa Transportista.");
	document.datos.tranro.focus();
	return;
}
if (document.datos.embnro.value == 0){
	alert("Debe ingresar el Nro. de Embarque.");
	document.datos.embnro.focus();
	return;
}

document.valida.location = "asignar_embarque_con_06.asp?tipo=<%= l_tipo%>&asiembnro="+document.datos.asiembnro.value +"&tarcod="+document.datos.tarcod.value + "&embnro="+document.datos.embnro.value + "&tranro="+document.datos.tranro.value + "&camnro="+document.camionero.datos.camnro.value;

}

function valido(){
	document.datos.submit();
}

function invalido(texto){
	alert(texto);
	document.datos.tarcod.focus();	
}

function CargarCamioneros(){
	document.camionero.location = "asignar_embarque_con_08.asp?embnro="+ document.datos.embnro.value;	
	document.datos.camcha.value = "";
	document.datos.camaco.value = "";	
	document.datos.tranro.value = 0;

}

</script>
<% 

Set l_rs = Server.CreateObject("ADODB.RecordSet")

select Case l_tipo
	Case "A":
		l_embnro  = 0
		l_camcha = ""
		l_camaco = ""
		l_tarcod = ""
		l_camnro = 0
		l_tranro = 0
		l_asiembobs = ""		
	Case "M":
		Set l_rs = Server.CreateObject("ADODB.RecordSet")
		l_asiembnro = request.querystring("cabnro")
		l_sql = "SELECT * "
		l_sql = l_sql & " FROM tkt_asiemb "
		l_sql  = l_sql  & " WHERE tkt_asiemb.asiembnro = " & l_asiembnro
		rsOpen l_rs, cn, l_sql, 0 
		if not l_rs.eof then
			l_embnro = l_rs("embnro")
			l_tarcod = l_rs("tarcod")
			l_camnro = l_rs("camnro")
			l_camcha = l_rs("camcha")
			l_camaco = l_rs("camaco")
			l_tranro = l_rs("tranro")
			l_asiembobs = l_rs("asiembobs")
		end if
		l_rs.Close
end select
%>
<body leftmargin="0" rightmargin="0" topmargin="0" bottommargin="0" onload="JavaScript:document.datos.tarcod.focus()">
<form name="datos" action="asignar_embarque_con_03.asp?tipo=<%= l_tipo %>" method="post" target="valida1">
<input type="hidden" name="asiembnro" value="<%= l_asiembnro %>">
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
					    <td align="right"><b>Tarjeta Nro:</b></td>
						<td>
							<input type="text" name="tarcod" size="8" maxlength="5" value="<%= l_tarcod %>">
						</td>
					</tr>
					
					<tr>
 				    	<td align="right" nowrap><b>Embarque:</b></td>
						<td colspan="3">
							<select name="embnro" size="1" style="width:300;" onchange="Javascript:CargarCamioneros();">
							<option value=0 selected>&laquo; Seleccione un Embarque Activo &raquo;</option>
							<%	l_sql = "SELECT embnro, embcod "
								l_sql  = l_sql  & " FROM tkt_embarque "
								
								rsOpen l_rs, cn, l_sql, 0
								do until l_rs.eof %>	
									<option value=<%= l_rs("embnro") %> > 
									<%= l_rs("embcod") %> </option>
									<%	l_rs.Movenext
								loop
								l_rs.Close %>
							</select>
							<script> document.datos.embnro.value="<%= l_embnro %>"</script>
						</td>
					</tr>					
					<tr>
					    <td align="right" nowrap><b>Camionero:</b></td>
					    <td align="left" >
						<iframe name="camionero" frameborder="0" width="100%" height="23" scrolling="No" src="asignar_embarque_con_08.asp?embnro=<%= l_embnro %>&camnro=<%= l_camnro %>"></iframe>					
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
					    <td align="right"><b>Observaciones:</b></td>
						<td>
							<TEXTAREA name="asiembobs" rows="3" cols="35" ><%= l_asiembobs %></TEXTAREA>
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

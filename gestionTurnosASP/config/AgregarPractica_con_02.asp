
<% Option Explicit %>
<!--#include virtual="/turnos/shared/db/conn_db.inc"-->
<% 

on error goto 0

'Datos del formulario

dim l_id
dim l_apellido
dim l_nombre  
dim l_dni     
dim l_domicilio
dim l_idobrasocial
'ADO
Dim l_tipo
Dim l_sql
Dim l_rs

dim l_idpractica 
dim l_idsolicitadapor
dim l_precio

Dim l_idvisita
Dim l_idpracticarealizada

Dim l_mediodepagoos
Dim l_idmediodepago
Dim l_osparticular

l_tipo = request.querystring("tipo")
l_idvisita = request("cabnro")

l_idpracticarealizada = request("idpracticarealizada")
l_idobrasocial=request("idobrasocial")

%>
<html>
<head>
<link href="/turnos/ess/shared/css/tables_gray.css" rel="StyleSheet" type="text/css">
<!--<link href="/turnos/shared/css/tables_gray.css" rel="StyleSheet" type="text/css">-->
<meta http-equiv="Content-Type" content="text/html; charset=iso-8859-1">
<title>Agregar Practica</title>
</head>
<script src="/turnos/shared/js/fn_valida.js"></script>
<script src="/turnos/shared/js/fn_fechas.js"></script>
<script src="/turnos/shared/js/fn_ayuda.js"></script>
<script src="/turnos/shared/js/fn_windows.js"></script>
<script src="/turnos/shared/js/fn_numeros.js"></script>
<script>
function Validar_Formulario(){



if (document.datos.practicaid.value == "0"){
	alert("Debe ingresar la Practica.");
	document.datos.practicaid.focus();
	return;
}

document.datos.precio2.value = document.datos.precio.value.replace(",", ".");
if (!validanumero(document.datos.precio2, 15, 4)){
		  alert("El Precio no es válido. Se permite hasta 15 enteros y 4 decimales.");	
		  document.datos.precio.focus();
		  document.datos.precio.select();
		  return;
}

<% if l_tipo = "A" then %>
if (document.datos.mediodepagoos.value == document.datos.idmediodepago.value)  {
	if (Trim(document.datos.idobrasocial.value) == "0"){
		alert("Debe ingresar la Obra Social.");
		document.datos.idobrasocial.focus();
		return;
	}
}

if (document.datos.importe.value == ""){
	alert("Debe ingresar un Importe mayor o igual a 0.");
	document.datos.importe.focus();
	return;
}

document.datos.importe2.value = document.datos.importe.value.replace(",", ".");

if (!validanumero(document.datos.importe2, 15, 4)){
		  alert("El Monto no es válido. Se permite hasta 15 enteros y 4 decimales.");	
		  document.datos.importe.focus();
		  document.datos.importe.select();
		  return;
}	

if (document.datos.importe.value != 0)  {
	if (Trim(document.datos.idmediodepago.value) == "0"){
		alert("Debe ingresar el Medio de Pago.");
		document.datos.idmediodepago.focus();
		return;
	}
}

if (document.datos.idmediodepago.value != "0")  {
	if (Trim(document.datos.importe.value) == "0"){
		alert("Debe ingresar el Importe.");
		document.datos.importe.focus();
		return;
	}
}
<% End If %>

valido();
}

function valido(){
	document.datos.submit();
}

function invalido(texto){
	alert(texto);
	document.datos.coudes.focus();
}


function ctrolmetodopago(){
	if (document.datos.mediodepagoos.value == document.datos.idmediodepago.value) {
			//document.datos.idobrasocial.readOnly = false;
			//document.datos.idobrasocial.className = 'habinp';			
			document.datos.idobrasocial.disabled = false;							
		}
		else {
			//document.datos.idobrasocial.readOnly = true;
			//document.datos.idobrasocial.className = 'deshabinp';		
			document.datos.idobrasocial.disabled = true;							
			document.datos.idobrasocial.value = 0;	
		}	

}


function Nuevo_Dialogo(w_in, pagina, ancho, alto)
{
 return w_in.showModalDialog(pagina,'', 'center:yes;dialogWidth:' + ancho.toString() + ';dialogHeight:' + alto.toString() + ';');
}
function Ayuda_Fecha(txt)
{
 var jsFecha = Nuevo_Dialogo(window, '/turnos/shared/js/calendar.html', 16, 15);

 if (jsFecha == null) txt.value = ''
 else txt.value = jsFecha;
}




function calcularprecio(){


	
	document.valida.location = "agregarpractica_con_06.asp?idos=" + document.datos.idos.value + "&practicaid="+ document.datos.practicaid.value ;	
}

function actualizarprecio(p_precio){	
	document.datos.precio.value = p_precio;
	

	// Si el medio de Pago es Obra social, copio el precio al importe
	if (document.datos.idmediodepago.value == document.datos.mediodepagoos.value ) { 
		document.datos.importe.value = p_precio;
	} 
 	else document.datos.importe.value = 0;
	
	

}	

</script>
<% 
Set l_rs = Server.CreateObject("ADODB.RecordSet")
'obtengo el Medio de Pago Obra Social
l_sql = "SELECT * "
l_sql = l_sql & " FROM mediosdepago "
l_sql  = l_sql  & " WHERE flag_obrasocial = -1 " 
l_sql = l_sql & " AND empnro = " & Session("empnro")
rsOpen l_rs, cn, l_sql, 0 
l_mediodepagoos = 0
if not l_rs.eof then
	l_mediodepagoos = l_rs("id")	
end if
l_rs.Close

'obtengo la Obra Social Particular
l_sql = "SELECT  * "
l_sql  = l_sql  & " FROM obrassociales "
l_sql  = l_sql  & " WHERE isnull(obrassociales.flag_particular,0) = -1 "	
l_sql = l_sql & " AND empnro = " & Session("empnro")								
rsOpen l_rs, cn, l_sql, 0 
l_osparticular = 0
if not l_rs.eof then
	l_osparticular = l_rs("id")	
end if
l_rs.Close

if l_idobrasocial = l_osparticular then
	l_idmediodepago = 0
	l_idobrasocial = l_osparticular
else
	l_idmediodepago = 	l_mediodepagoos
end if


select Case l_tipo
	Case "A":
			l_idpractica = 0
			l_idsolicitadapor = 0
			l_precio = 0
	Case "M":
		
		l_id = request.querystring("cabnro")
		l_sql = "SELECT * "
		l_sql = l_sql & " FROM practicasrealizadas "
		l_sql  = l_sql  & " WHERE id = " & l_idpracticarealizada
		rsOpen l_rs, cn, l_sql, 0 
		if not l_rs.eof then
			l_idpractica = l_rs("idpractica")
			l_idsolicitadapor = l_rs("idsolicitadapor") 
			l_precio = l_rs("precio")
		end if
		l_rs.Close
end select
%>
<body leftmargin="0" rightmargin="0" topmargin="0" bottommargin="0" onload="javascript:document.datos.motivo.focus();">
<form name="datos" action="AgregarPractica_con_03.asp?tipo=<%= l_tipo %>" method="post" target="valida">
<input type="hidden" name="idvisita" value="<%= l_idvisita %>">
<input type="hidden" name="idpracticarealizada" value="<%= l_idpracticarealizada %>">
<input type="hidden" name="idos" value="<%= l_idobrasocial %>">


<input type="hidden" name="mediodepagoos" value="<%= l_mediodepagoos %>">
<input type="hidden" name="osparticular" value="<%= l_osparticular %>">

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
				<td>
					<table cellspacing="0" cellpadding="0" border="0">
					
											
					<tr>
						<td  align="right" nowrap><b>Practica (*): </b></td>
						<td colspan="3"><select name="practicaid" size="1" style="width:200;" onchange="calcularprecio();">
								<option value=0 selected>Seleccione una Practica</option>
								<%Set l_rs = Server.CreateObject("ADODB.RecordSet")
								l_sql = "SELECT  * "
								l_sql  = l_sql  & " FROM practicas "
								l_sql = l_sql & " where empnro = " & Session("empnro")
								l_sql  = l_sql  & " ORDER BY descripcion "
								rsOpen l_rs, cn, l_sql, 0
								do until l_rs.eof		%>	
								<option value= <%= l_rs("id") %> > 
								<%= l_rs("descripcion") %> </option>
								<%	l_rs.Movenext
								loop
								l_rs.Close %>
							</select>
							<script>document.datos.practicaid.value="<%= l_idpractica %>"</script>
						</td>					
					</tr>	
					
					<tr>
						<td  align="right" nowrap><b>Solicitado por : </b></td>
						<td colspan="3"><select name="idrecursoreservable" size="1" style="width:200;">
								<option value=0 selected>Ningun Profesional</option>
								<%Set l_rs = Server.CreateObject("ADODB.RecordSet")
								l_sql = "SELECT  * "
								l_sql  = l_sql  & " FROM recursosreservables "
								l_sql = l_sql & " where empnro = " & Session("empnro")
								l_sql  = l_sql  & " ORDER BY descripcion "
								rsOpen l_rs, cn, l_sql, 0
								do until l_rs.eof		%>	
								<option value= <%= l_rs("id") %> > 
								<%= l_rs("descripcion") %> </option>
								<%	l_rs.Movenext
								loop
								l_rs.Close %>
							</select>
							<script>document.datos.idrecursoreservable.value="<%= l_idsolicitadapor %>"</script>							
						</td>					
					</tr>		

					<tr>
					    <td align="right"><b>Precio:</b></td>
						<td colspan="3">
							<input align="right" type="text" name="precio" size="20" maxlength="20" value="<%= l_precio %>">
							<input type="hidden" name="precio2" value="">							
						</td>
					</tr>		
					<% if l_tipo = "A" then %>	
					<tr>
					    
						<td colspan="4">
							&nbsp;						
						</td>
					</tr>						

					<tr>
						<td  align="right" nowrap><b>Medio de Pago: </b></td>
						<td colspan="3"><select name="idmediodepago" size="1" style="width:200;" onchange="ctrolmetodopago();">
								<option value=0 selected>Seleccione un Medio</option>
								<%Set l_rs = Server.CreateObject("ADODB.RecordSet")
								l_sql = "SELECT  * "
								l_sql  = l_sql  & " FROM mediosdepago "
								l_sql = l_sql & " where empnro = " & Session("empnro")
								l_sql  = l_sql  & " ORDER BY titulo "
								rsOpen l_rs, cn, l_sql, 0
								do until l_rs.eof		%>	
								<option value= <%= l_rs("id") %> > 
								<%= l_rs("titulo") %> </option>
								<%	l_rs.Movenext
								loop
								l_rs.Close %>
							</select>
							<script>document.datos.idmediodepago.value="<%= l_idmediodepago %>"</script>

						</td>					
					</tr>		
					<tr>
						<td  align="right" nowrap><b>Obra Social: </b></td>
						<td colspan="3"><select name="idobrasocial" size="1" style="width:200;">
								<option value=0 selected>Seleccione una OS</option>
								<%Set l_rs = Server.CreateObject("ADODB.RecordSet")
								l_sql = "SELECT  * "
								l_sql  = l_sql  & " FROM obrassociales "
								l_sql  = l_sql  & " WHERE isnull(obrassociales.flag_particular,0) = 0 "	
								l_sql = l_sql & " AND empnro = " & Session("empnro")								
								l_sql  = l_sql  & " ORDER BY descripcion "
								rsOpen l_rs, cn, l_sql, 0
								do until l_rs.eof		%>	
								<option value= <%= l_rs("id") %> > 
								<%= l_rs("descripcion") %> </option>
								<%	l_rs.Movenext
								loop
								l_rs.Close %>
							</select>
							<script>document.datos.idobrasocial.value="<%= l_idobrasocial %>"</script>
							<script>ctrolmetodopago();</script>
						</td>					
					</tr>		
					<tr>
					    <td align="right"><b>Nro:</b></td>
						<td>
							<input   type="text" name="nro" size="20" maxlength="20" value="<%'= l_nro %>">
						</td>					
					</tr>		
					<tr>
					    <td align="right"><b>Importe:</b></td>
						<td>
							<input align="right" type="text" name="importe" size="20" maxlength="20" value="0">
							<input type="hidden" name="importe2" value="">
						</td>					
					</tr>												
					
					<% End If %>
					
					</table>
				</td>
			</tr>
		</table>
	</td>
</tr>
<tr>
    <td colspan="2" align="right" class="th">
		<a class=sidebtnABM href="Javascript:Validar_Formulario()">Aceptar</a>
		<a class=sidebtnABM href="Javascript:window.close()">Cancelar</a>
	</td>
</tr>

</table>
<iframe name="valida" style="visibility=hidden;" src="" width="100%" height="100%"></iframe> 
</form>
<%
set l_rs = nothing
cn.Close
set cn = nothing
%>
</body>
</html>

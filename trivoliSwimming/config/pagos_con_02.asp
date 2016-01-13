<% Option Explicit %>
<!--#include virtual="/turnos/shared/db/conn_db.inc"-->
<% 
'Archivo: companies_con_02.asp
'Descripción: ABM de Companies
'Autor : Raul Chinestra
'Fecha: 26/11/2007

'Datos del formulario
Dim l_id
Dim l_titulo
Dim l_fecha

'ADO
Dim l_tipo
Dim l_sql
Dim l_rs

Dim l_flag_activo
Dim l_idpracticarealizada


Dim l_idmediodepago

Dim l_idobrasocial
Dim l_nro
Dim l_importe
Dim l_mediodepagoos 
Dim l_idospaciente
Dim l_flag_particular


l_tipo = request.querystring("tipo")
l_idpracticarealizada = request.querystring("idpracticarealizada")


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


'obtengo la Obra Social del paciente
l_sql = "select isnull(clientespacientes.idobrasocial, 0) idobrasocial,isnull(obrassociales.flag_particular, 0) flag_particular "
l_sql = l_sql & " from practicasrealizadas "
l_sql = l_sql & " inner join visitas on practicasrealizadas.idvisita = visitas.id "
l_sql = l_sql & " inner join clientespacientes on clientespacientes.id = visitas.idpaciente "
l_sql = l_sql & " inner join obrassociales on obrassociales.id = clientespacientes.idobrasocial "
l_sql  = l_sql  & " WHERE practicasrealizadas.id =  " &l_idpracticarealizada
l_sql = l_sql & " AND practicasrealizadas.empnro = " & Session("empnro")

'response.write l_sql&"</BR>"

rsOpen l_rs, cn, l_sql, 0 
l_idospaciente = 0
l_flag_particular = 0
if not l_rs.eof then
	l_idospaciente = l_rs("idobrasocial")	
	l_flag_particular = l_rs("flag_particular")		
end if
l_rs.Close



%>
<html>
<head>
<link href="/turnos/ess/shared/css/tables_gray.css" rel="StyleSheet" type="text/css">
<meta http-equiv="Content-Type" content="text/html; charset=iso-8859-1">
<title>Detalle de Pagos</title>
</head>
<script src="/turnos/shared/js/fn_valida.js"></script>
<script src="/turnos/shared/js/fn_fechas.js"></script>
<script src="/turnos/shared/js/fn_ayuda.js"></script>
<script src="/turnos/shared/js/fn_windows.js"></script>
<script src="/turnos/shared/js/fn_numeros.js"></script>

<!-- Comienzo Datepicker -->
<link rel="stylesheet" href="../js/themes/smoothness/jquery-ui.css">
<script src="../js/jquery-1.8.0.js"></script>
<script src="../js/jquery-ui.js"></script>  
<script src="../js/jquery.ui.datepicker-es.js"></script>
<script>
$(function () {
$.datepicker.setDefaults($.datepicker.regional["es"]);
$("#datepicker").datepicker({
firstDay: 1
});

		
$( "#fecha" ).datepicker({
	showOn: "button",
	buttonImage: "/turnos/shared/images/calendar1.png",
	buttonImageOnly: true
});


});
</script>
<!-- Final Datepicker -->

<script>
function Validar_Formulario(){

if ((document.datos.fecha.value == "")&&(!validarfecha(document.datos.fecha))){
	 document.datos.fecha.focus();
	 return;
}

if (Trim(document.datos.idmediodepago.value) == "0"){
	alert("Debe ingresar el Medio de Pago.");
	document.datos.idmediodepago.focus();
	return;
}
if (document.datos.mediodepagoos.value == document.datos.idmediodepago.value)  {
	if (Trim(document.datos.idobrasocial.value) == "0"){
		alert("Debe ingresar la Obra Social.");
		document.datos.idobrasocial.focus();
		return;
	}
}

document.datos.importe2.value = document.datos.importe.value.replace(",", ".");
  
if (!validanumero(document.datos.importe2, 15, 4)){
		  alert("El Monto no es válido. Se permite hasta 15 enteros y 4 decimales.");	
		  document.datos.importe.focus();
		  document.datos.importe.select();
		  return;
}	



valido();

}

function valido(){
	document.datos.submit();
}

function invalido(texto){
	alert(texto);
	document.datos.agedes.focus();
}

function Ayuda_Fecha(txt)
{
 var jsFecha = Nuevo_Dialogo(window, '/turnos/shared/js/calendar.html', 16, 15);

 if (jsFecha == null) txt.value = ''
 else txt.value = jsFecha;
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
			//document.datos.idobrasocial.value = 0;	
		}	

}

</script>
<% 
select Case l_tipo
	Case "A":
	
		l_fecha = Date()
		l_nro = ""
		l_importe = "0"

		if l_flag_particular = 0 then
			l_idmediodepago = l_mediodepagoos
			l_idobrasocial = l_idospaciente			
		else
			l_idmediodepago = "0"
			l_idobrasocial = "0"		
		end if
		
	Case "M":
		'Set l_rs = Server.CreateObject("ADODB.RecordSet")
		l_id = request.querystring("cabnro")
		l_sql = "SELECT * "
		l_sql = l_sql & " FROM pagos "
		l_sql  = l_sql  & " WHERE id = " & l_id
		rsOpen l_rs, cn, l_sql, 0 
		if not l_rs.eof then
			l_idmediodepago = l_rs("idmediodepago")
			l_fecha = l_rs("fecha")
			l_idobrasocial = l_rs("idobrasocial")
			l_nro = l_rs("nro")
			l_importe = l_rs("importe")
	
		end if
		l_rs.Close
end select
%>
<body leftmargin="0" rightmargin="0" topmargin="0" bottommargin="0" onload="JavaScript:document.datos.fecha.focus()">
<form name="datos" action="pagos_con_03.asp?tipo=<%= l_tipo %>" method="post" target="valida">
<input type="Hidden" name="id" value="<%= l_id %>">
<input type="Hidden" name="idpracticarealizada" value="<%= l_idpracticarealizada %>">
<input type="Hidden" name="mediodepagoos" value="<%= l_mediodepagoos %>">

<table cellspacing="0" cellpadding="0" border="0" width="100%" height="100%">
<tr>
    <td class="th2" nowrap>Detalle de Pago</td>
	<td class="th2" align="right">
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
					    <td align="right" nowrap width="0"><b>Fecha:</b></td>
						<td align="left" nowrap width="0" >
						    <input type="text"  id="fecha" name="fecha" size="10" maxlength="10" value="<%= l_fecha %>">							
						</td>																	
					</tr>							
					<tr>
						<td  align="right" nowrap><b>Medio de Pago: </b></td>
						<td colspan="3"><select name="idmediodepago" size="1" style="width:200;" onchange="ctrolmetodopago();">
								<option value=0 selected>Seleccione Medio de Pago</option>
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
							<input   type="text" name="nro" size="20" maxlength="20" value="<%= l_nro %>">
						</td>					
					</tr>		
					<tr>
					    <td align="right"><b>Importe:</b></td>
						<td>
							<input align="right" type="text" name="importe" size="20" maxlength="20" value="<%= l_importe %>">
							<input type="hidden" name="importe2" value="">
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

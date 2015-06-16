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
Dim l_descripcion

'ADO
Dim l_tipo
Dim l_sql
Dim l_rs


Dim l_horainicial 
Dim l_horafinal
Dim l_intervaloturnominutos
Dim l_calfec

Dim l_idrecursoreservable




l_tipo = request.querystring("tipo")
l_idrecursoreservable = request.querystring("idrecursoreservable")
l_calfec  = request.querystring("fechadesde")


%>
<html>
<head>
<link href="/turnos/ess/shared/css/tables_gray.css" rel="StyleSheet" type="text/css">
<meta http-equiv="Content-Type" content="text/html; charset=iso-8859-1">
<title>Agregar Visitas sin Turno</title>
</head>
<script src="/turnos/shared/js/fn_ayuda.js"></script>
<script src="/turnos/shared/js/fn_windows.js"></script>
<script src="/turnos/shared/js/fn_valida.js"></script>
<script src="/turnos/shared/js/fn_fechas.js"></script>
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

		
$( "#calfec" ).datepicker({
	showOn: "button",
	buttonImage: "/turnos/shared/images/calendar1.png",
	buttonImageOnly: true
});



});
</script>
<!-- Final Datepicker -->
<script>
function Validar_Formulario(){


if ((document.datos.calfec.value == "")&&(!validarfecha(document.datos.calfec))){
	 document.datos.calfec.focus();
	 return;
}

if (document.datos.pacienteid.value == "0"){
	alert("Debe ingresar el Paciente.");
	document.datos.pacienteid.focus();
	return;
}

/*
if (Trim(document.datos.titulo.value) == ""){
	alert("Debe ingresar el T&iacute;tulo.");
	document.datos.titulo.focus();
	return;
}


if (Trim(document.datos.descripcion.value) == ""){
	alert("Debe ingresar la Descripción.");
	document.datos.descripcion.focus();
	return;
}
/*
if (!stringValido(document.datos.agedes.value)){
	alert("La Descripción contiene caracteres inválidos.");
	document.datos.agedes.focus();
	return;
}
*/
var d=document.datos;
document.valida.location = "altavisita_con_06.asp?pacienteid="+document.datos.pacienteid.value ; 


//valido();

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

function EncontrePaciente(id, apellido, nombre, nrohistoriaclinica, dni, domicilio, tel, osid, os){
	document.datos.pacienteid.value = id;
	document.datos.apellido.value = apellido;
	document.datos.nombre.value = nombre;
	document.datos.nrohistoriaclinica.value = nrohistoriaclinica;
	document.datos.dni.value = dni;
	//document.datos.domicilio.value = domicilio;
	//document.datos.tel.value = tel;
	//document.datos.osid.value = osid;
	//document.datos.os.value = os;
	//document.datos.coudes.focus();
}

function BuscarPaciente(){
	abrirVentana('Buscarpacientes_con_00.asp?Tipo=A','',600,250);
}

</script>
<% 
select Case l_tipo
	Case "A":
		l_titulo = ""
		l_descripcion = ""
	Case "M":
		'Set l_rs = Server.CreateObject("ADODB.RecordSet")
		'l_id = request.querystring("cabnro")
		'l_sql = "SELECT * "
		'l_sql = l_sql & " FROM templatereservasdetalleresumido "
		'l_sql  = l_sql  & " WHERE id = " & l_id
		'rsOpen l_rs, cn, l_sql, 0 
		'if not l_rs.eof then
		'	l_titulo = l_rs("titulo")
		'	l_horainicial = l_rs("horainicial") 
		'	l_horafinal = l_rs("horafinal") 
		'	l_intervaloturnominutos = l_rs("intervaloturnominutos") 
		'	l_do =  l_rs("dia1") 
		'	l_lu =  l_rs("dia2")
		'	l_ma =  l_rs("dia3")
		'	l_mi =  l_rs("dia4")
		'	l_ju =  l_rs("dia5")
		'	l_vi =  l_rs("dia6")
		'	l_sa =  l_rs("dia7")
		'end if
		'l_rs.Close
end select
%>
<body leftmargin="0" rightmargin="0" topmargin="0" bottommargin="0" onload="JavaScript:document.datos.titulo.focus()">
<form name="datos" action="altavisita_con_03.asp?tipo=<%= l_tipo %>" method="post" target="valida">
<input type="Hidden" name="idrecursoreservable" value="<%= l_idrecursoreservable %>">


<table cellspacing="0" cellpadding="0" border="0" width="100%" height="50%">
<tr>
    <td class="th2" nowrap>&nbsp;</td>
</tr>
<tr>
	<td >
		<table border="0" cellspacing="0" cellpadding="0">
			<tr>
				<td width="50%"></td>
				<td>
					<table cellspacing="0" cellpadding="0" border="0">
					<tr>
					    <td align="right" nowrap width="0"><b>Fecha (*):</b></td>
						<td align="left" nowrap width="0" >
						    <input type="text" id="calfec" name="calfec" size="10" maxlength="10" value="<%= l_calfec %>">							
						</td>																	
					</tr>	
					<tr>
					    <td align="right" nowrap width="0"><b>Paciente (*):</b></td>
						<td align="left" nowrap width="0" >
						    <input type="hidden" name="pacienteid" size="10" maxlength="10" value="0">
							<a href="Javascript:BuscarPaciente();"><img src="/turnos/shared/images/BuscarPaciente.png" border="0" alt="Buscar Paciente"></a>	
						</td>																	
					</tr>
					<tr>
					    <td align="right"><b>Apellido:</b></td>
						<td>
							<input class="deshabinp" readonly="" type="text" name="apellido" size="20" maxlength="20" value="<%'= l_apellido %>">							
						</td>
					    <td align="right"><b>Nombre:</b></td>						
						<td>
							<input class="deshabinp" readonly="" type="text" name="nombre" size="20" maxlength="20" value="<%'= l_nombre %>">
						</td>						
					</tr>					
					<tr>
					    <td align="right"><b>D.N.I.:</b></td>
						<td>
							<input class="deshabinp" readonly="" type="text" name="dni" size="20" maxlength="20" value="<%'= l_dni %>">
						</td>
					    <td align="right"><b>Nro. Historia Cl&iacute;nica:</b></td>
						<td>
							<input class="deshabinp" readonly="" type="text" name="nrohistoriaclinica" size="20" maxlength="20" value="<%'= l_nrohistoriaclinica %>">
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
<iframe name="valida"  style="visibility=hidden;" src="" width="100%" height="100%"></iframe> 
</form>
<%
set l_rs = nothing
cn.Close
set cn = nothing
%>
</body>
</html>

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

Dim l_hd
Dim l_md




l_tipo = request.querystring("tipo")
l_id = request.querystring("id")


%>
<html>
<head>
<link href="/turnos/ess/shared/css/tables_gray.css" rel="StyleSheet" type="text/css">
<meta http-equiv="Content-Type" content="text/html; charset=iso-8859-1">
<title>Detalle de Calendarios</title>
</head>
<script src="/turnos/shared/js/fn_ayuda.js"></script>
<script src="/turnos/shared/js/fn_windows.js"></script>
<script src="/turnos/shared/js/fn_valida.js"></script>
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

if (document.datos.calfec.value == ""){
	alert("Debe ingresar la fecha.");
	document.datos.calfec.focus();
	return;
}


if (document.datos.calhordes1.value == "" || document.datos.calhordes2.value == ""){
	alert("Debe ingresar la Hora Desde");
	if (document.datos.calhordes1.value == "") {document.datos.calhordes1.focus();}
	else if (document.datos.calhordes2.value == "") {document.datos.calhordes2.focus();}
	return;
}

	
if (document.datos.calhorhas1.value == "" || document.datos.calhorhas2.value == ""){
	alert("Debe ingresar la Hora Hasta");
	if (document.datos.calhorhas1.value == "") {document.datos.calhorhas1.focus();}
	else if (document.datos.calhorhas2.value == "") {document.datos.calhorhas2.focus();}
	return;
}
	
if (document.datos.intervaloTurnoMinutos.value == ""){
	alert("Debe ingresar el intervalo de los turnos");
	document.datos.intervaloTurnoMinutos.focus();	
	return;
}


if (isNaN(document.datos.intervaloTurnoMinutos.value)) {
	alert("El intervalo de minutos debe ser numerico.");
	document.datos.intervaloTurnoMinutos.focus();
	return;
}

if (document.datos.calhorhas1.value < document.datos.calhordes1.value) {
	alert("La hora desde debe ser menor a la hora hasta");
	document.datos.calhorhas1.focus();	
	return;
} else if (document.datos.calhorhas1.value == document.datos.calhordes1.value && document.datos.calhorhas2.value <= document.datos.calhordes2.value) {
	alert("La hora (y minutos) desde debe ser menor a la hora (y minutos) hasta");
	document.datos.calhorhas2.focus();	
	return;
}

var d=document.datos;
document.valida.location = "calendarios_con_06.asp?id=<%= l_id%>&calfec="+document.datos.calfec.value + "&calhordes1="+document.datos.calhordes1.value + "&calhordes2="+document.datos.calhordes2.value + "&calhorhas1="+document.datos.calhorhas1.value + "&calhorhas2="+document.datos.calhorhas2.value + "&intervaloTurnoMinutos="+document.datos.intervaloTurnoMinutos.value ; 


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
<form name="datos" action="calendarios_con_03.asp?tipo=<%= l_tipo %>" method="post" target="valida">
<input type="Hidden" name="id" value="<%= l_id %>">


<table cellspacing="0" cellpadding="0" border="0" width="100%" height="50%">
<tr>
    <td class="th2" nowrap>Detalle de Calendarios</td>
</tr>
<tr>
	<td >
		<table border="0" cellspacing="0" cellpadding="0">
			<tr>
				<td width="50%"></td>
				<td>
					<table cellspacing="0" cellpadding="0" border="0">
					<tr>
					    <td align="right" nowrap width="0"><b>Fecha:</b></td>
						<td align="left" nowrap width="0" >
						    <input type="text" id="calfec" name="calfec" size="10" maxlength="10" value="<%= l_calfec %>">							
						</td>																	
					</tr>	
					<!-- 					
					<tr>
					    <td align="right"><b>Descripción:</b></td>
						<td>
							<input type="text" name="descripcion" size="40" maxlength="50" value="<%'= l_descripcion %>">
						</td>
					</tr>  -->			
					</table>
				</td>
				<td width="50%"></td>
			</tr>
		</table>
	</td>
</tr>



<tr>
   <td >
		<table border="0" cellspacing="0" cellpadding="0">
			<tr>
	<td align="right"><b>Hora Desde:</b></td>
	<td>			
			<select name="calhordes1" size="1" style="width:50;">
				<%
				l_hd = 0  
				do while clng(l_hd) < 24 %>
				<option value= <%= right("0" & l_hd, 2) %>> <%= right("0" & l_hd, 2) %> </option>
				<%	l_hd = clng(l_hd) + 1
				loop
				%>
			</select>
			<script>document.datos.calhordes1.value="<%= left(l_horainicial,2) %>"</script>				
		<b>:</b>			
			<select name="calhordes2" size="1" style="width:50;">
				<%
				l_md = 0  
				do while clng(l_md) < 60 %>
				<option value= <%= right("0" & l_md, 2) %>> <%= right("0" & l_md, 2) %> </option>
				<%	l_md = clng(l_md) + 15
				loop
				%>
			</select>	
			<script>document.datos.calhordes2.value="<%= right(l_horainicial,2) %>"</script>			
	</td>
	<td align="right"><b>Hora Hasta:</b></td>
	<td>			
			<select name="calhorhas1" size="1" style="width:50;">
				<%
				l_hd = 0  
				do while clng(l_hd) < 24 %>
				<option value= <%= right("0" & l_hd, 2) %>> <%= right("0" & l_hd, 2) %> </option>
				<%	l_hd = clng(l_hd) + 1
				loop
				%>
			</select>	
			<script>document.datos.calhorhas1.value="<%= left(l_horafinal,2) %>"</script>				
		<b>:</b>			
			<select name="calhorhas2" size="1" style="width:50;">
				<%
				l_md = 0  
				do while clng(l_md) < 60 %>
				<option value= <%= right("0" & l_md, 2) %>> <%= right("0" & l_md, 2) %> </option>
				<%	l_md = clng(l_md) + 15
				loop
				%>
			</select>				
			<script>document.datos.calhorhas2.value="<%= right(l_horafinal,2) %>"</script>				
	</td>
		</tr>
		</table>
	</td>	
</tr>

<tr>
   <td >
		<table border="0" cellspacing="0" cellpadding="0">
					<tr>
					    <td align="right"><b>Intervalo Minutos:</b></td>
						<td>
							<input type="text" name="intervaloTurnoMinutos" size="10" maxlength="10" value="<%= l_intervaloTurnoMinutos %>">
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
<iframe name="valida"  style="visibility=hidden;" src="" width="100%" height="100%"></iframe> 
</form>
<%
set l_rs = nothing
cn.Close
set cn = nothing
%>
</body>
</html>

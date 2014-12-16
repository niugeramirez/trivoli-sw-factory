<%Option Explicit %>
<!--#include virtual="/serviciolocal/ess/shared/inc/const.inc"-->
<!--#include virtual="/serviciolocal/shared/inc/sec.inc"-->
<!--#include virtual="/serviciolocal/shared/inc/const.inc"-->
<!--#include virtual="/serviciolocal/shared/db/conn_db.inc"-->
<!--#include virtual="/serviciolocal/shared/inc/fecha.inc"-->
<!--#include virtual="/serviciolocal/shared/inc/sqls.inc"-->
<!--#include virtual="/serviciolocal/ess/shared/inc/accesoESS.inc"-->
<% 
'---------------------------------------------------------------------------------------
'Archivo	: pedido_emp_vac_gti_02.asp
'Descripción: pedido vacaciones
'Autor: Scarpa D.
'Fecha: 08/10/2004
'Modificado:  11/08/2006 - Moro L. - Consulta para mostrar el nombre del empleado
'			  30-07-2007 - Diego Rosso - Se agrego src=blanc.asp por https.
'---------------------------------------------------------------------------------------

' Variables
on error goto 0

Dim l_ternro
Dim l_terapenom
Dim l_vdiapednro

Dim l_vacnro		 
Dim l_vdiapeddesde	 
Dim l_vdiapedhasta	 
Dim l_vdiapedcant    
Dim l_vdiaspedestado 
Dim l_vdiaspedhabiles 
Dim l_vdiaspedferiados 
Dim l_vdiaspednohabiles

Dim l_diasyaped
Dim l_diascorridos

Dim l_tipvacnro
Dim l_tipvacdesabr

dim l_tipo

dim l_seguir
dim l_rs
dim l_rs1
dim l_rs2
dim l_sql

dim l_deshabilitar

Set l_rs  = Server.CreateObject("ADODB.RecordSet")
Set l_rs1 = Server.CreateObject("ADODB.RecordSet")
Set l_rs2 = Server.CreateObject("ADODB.RecordSet")			

l_tipo          = Request.QueryString("tipo")
l_vdiapednro    = Request.QueryString("vdiapednro")

'l_ternro = l_ess_ternro
dim leg
leg = Session("empleg")
if leg = "" then
    response.write "NO SE HA SELECCIONADO UN EMPLEADO<BR>"
	Response.End
end if

l_sql = "SELECT ternro, ternom, ternom2, terape, terape2 FROM empleado WHERE empleado.empleg = " & leg
l_rs.Open l_sql, cn
if l_rs.eof then
    response.write "NO SE HA SELECCIONADO UN EMPLEADO<BR>"
	response.end
else 
  l_ternro = l_rs("ternro")
  l_terapenom = l_rs("terape") & " " & l_rs("terape2") & ", " & l_rs("ternom") & " " & l_rs("ternom2")
end if
l_rs.close

'inicializo el valor
l_vacnro		= 0 
	
select Case l_tipo
	Case "A":
			l_vacnro		 = ""
			l_vdiapeddesde	 = Date()
			l_vdiapedhasta	 = ""
			l_vdiapedcant    = 1
			l_vdiaspedestado = 0
			l_vdiaspedhabiles = ""
			l_vdiaspedferiados = ""
			l_vdiaspednohabiles = ""
			l_deshabilitar = false
	Case "M":
		If len(trim(l_vdiapednro)) = 0 then
			response.write("<script>alert('Debe seleccionar un periodo de vacaciones');window.close();</script>")
		end if
		
		l_sql = "SELECT vacnro, vdiapednro,  "  
		l_sql = l_sql & " vdiapeddesde,  "
		l_sql = l_sql & " vdiapedhasta,  "
		l_sql = l_sql & " vdiapedcant,   "
		l_sql = l_sql & " vdiaspedestado,  "
		l_sql = l_sql & " vdiaspedhabiles,   "
		l_sql = l_sql & " vdiaspednohabiles,   "
		l_sql = l_sql & " vdiaspedferiados   "
		l_sql = l_sql & " FROM  vacdiasped"
		l_sql = l_sql & " WHERE vacdiasped.vdiapednro = " & l_vdiapednro
		l_sql = l_sql & " AND   vacdiasped.ternro     = " & l_ternro
		
		rsOpen l_rs, cn, l_sql, 0 
		if not l_rs.eof then
			l_vacnro			= l_rs("vacnro")
			l_vdiapeddesde		= l_rs("vdiapeddesde")
			l_vdiapedhasta		= l_rs("vdiapedhasta")
			'l_vdiapedcant		= l_rs("vdiapedcant")
			l_vdiapedcant		= l_rs("vdiaspedhabiles")
			l_vdiaspedestado	= l_rs("vdiaspedestado")
			l_vdiaspedhabiles	= l_rs("vdiaspedhabiles")
			l_vdiaspednohabiles	= l_rs("vdiaspednohabiles")
			l_vdiaspedferiados	= l_rs("vdiaspedferiados")
		end if
		l_rs.Close
		
		'Controlo si la licencia se puede modificar
		
		if CInt(l_vdiaspedestado) = -1 then
		
	        l_sql = "SELECT elfechadesde,elfechahasta, elcantdias, vacnotifestado "
	        l_sql = l_sql & "FROM emp_lic INNER JOIN lic_vacacion ON lic_vacacion.emp_licnro = emp_lic.emp_licnro "
	        l_sql = l_sql & "LEFT JOIN vacnotif ON vacnotif.emp_licnro = emp_lic.emp_licnro "
	        l_sql = l_sql & "WHERE licestnro=2 AND empleado = " & l_ternro & " and emp_lic.tdnro = 2 AND lic_vacacion.vacnro = " & l_vacnro
	        l_sql = l_sql & " AND elfechadesde >= " & cambiafecha(l_vdiapeddesde,"YMD",true)
	        l_sql = l_sql & " AND elfechahasta <= " & cambiafecha(l_vdiapedhasta,"YMD",true)
			
		    rsOpen l_rs, cn, l_sql, 0 
	
			l_deshabilitar	= (not l_rs.eof)
	
			l_rs.Close

		else

			l_deshabilitar	= false

		end if

end select
%>

<%
response.write "<script languaje='javascript'>" & vbCrLf

response.write " function CargarFechas(){ " & vbCrLf

	l_sql = "SELECT vacacion.vacnro, vacacion.vacfecdesde, vacacion.vacfechasta  "  
	l_sql = l_sql & " FROM  vacacion"
	rsOpen l_rs1, cn, l_sql, 0 
	
	response.write " document.datos.vacfecdesde.value = ''; " & vbCrLf
	response.write " document.datos.vacfechasta.value = ''; " & vbCrLf
	
	do while NOT l_rs1.EOF 
		response.write "if (document.datos.vacnro.value == "& l_rs1(0) &") { " & vbCrLf
		response.write " document.datos.vacfecdesde.value = '" & l_rs1(1) & "'; " & vbCrLf
		response.write " document.datos.vacfechasta.value = '" & l_rs1(2) & "'; " & vbCrLf
		response.write " };" & vbCrLf
		l_rs1.MoveNext
	loop	
	l_rs1.Close
	response.write " CalculoDias(); " & vbCrLf

response.write "};" & vbCrLf


response.write "function CalculoDias(){ " & vbCrLf

	l_sql = "SELECT vacdiascor.vacnro, vacdiascor.tipvacnro, "
	l_sql = l_sql & " vacdiascor.vdiascorcant "
	l_sql = l_sql & " FROM  vacdiascor "
	l_sql = l_sql & " WHERE vacdiascor.ternro =  " & l_ternro
	response.write  "/*" & l_sql & "*/"
	rsOpen l_rs1, cn, l_sql, 0 

	response.write "document.datos.tipvacnro.value = '' ;" & vbCrLf
	response.write "document.datos.tipvacdesabr.value = '';" & vbCrLf
	response.write "document.datos.vdiascorcant.value = '0';" & vbCrLf
			
	do while NOT l_rs1.EOF 
		response.write "if (document.datos.vacnro.value == "& l_rs1(0) &") { " & vbCrLf
			response.write "document.datos.vdiascorcant.value = '" & l_rs1(2) & "';" & vbCrLf

			l_sql = "SELECT tipovacac.tipvacnro, tipovacac.tipvacdesabr  "
			l_sql = l_sql & "FROM  tipovacac "
			l_sql = l_sql & "WHERE tipovacac.tipvacnro =  " & l_rs1(1)
			rsOpen l_rs2, cn, l_sql, 0 
			if NOT l_rs2.EOF then
				response.write "document.datos.tipvacnro.value = " & l_rs2(0) & ";" & vbCrLf
				response.write "document.datos.tipvacdesabr.value = '" & l_rs2(1) & "';" & vbCrLf
			end if
			l_rs2.Close

		response.write "};" & vbCrLf
		l_rs1.MoveNext
	loop
	
	l_rs1.close
	
	' Calcular ya pedidos -----------------------------------------------------

	l_sql = "SELECT vacdiasped.vacnro, vacdiasped.ternro, SUM(vacdiasped.vdiaspedhabiles) as suma"
	l_sql = l_sql & " FROM  vacdiasped "
	l_sql = l_sql & " WHERE vacdiasped.ternro =  " & l_ternro
	l_sql = l_sql & "   AND vacdiasped.vdiaspedestado = -1 " 
	l_sql = l_sql & " GROUP BY vacdiasped.ternro, vacdiasped.vacnro " 
	rsOpen l_rs1, cn, l_sql, 0 
	
	response.write "    document.datos.diasyaped.value = 0 " & vbCrLf
	
	do while NOT l_rs1.EOF 
		response.write " if (document.datos.vacnro.value == "& l_rs1(0) &") { " & vbCrLf
		response.write "    document.datos.diasyaped.value = " & l_rs1("suma") & vbCrLf
'		response.write "	document.datos.diasyaped.value = document.datos.diasyaped.value - document.datos.vdiapedcant.value;" 
		response.write "};" & vbCrLf
    	l_rs1.MoveNext
	loop

	response.write "document.datos.diaspendientes.value = parseInt(document.datos.vdiascorcant.value) - parseInt(document.datos.diasyaped.value);" & vbCrLf
	 
response.write "};" & vbCrLf
response.write "</script>" & vbCrLf

l_rs1.Close
set l_rs1 = nothing

%>

<html>
<head>
<link href="../<%= c_estilo %>" rel="StyleSheet" type="text/css">
<meta http-equiv="Content-Type" content="text/html; charset=iso-8859-1">
<title>Pedidos de Vacaciones</title>
<script src="/serviciolocal/shared/js/fn_windows.js"></script>
<script src="/serviciolocal/shared/js/fn_confirm.js"></script>
<script src="/serviciolocal/shared/js/fn_hora.js"></script>
<script src="/serviciolocal/shared/js/fn_fechas.js"></script>
<script src="/serviciolocal/shared/js/fn_ayuda.js"></script>
<script src="/serviciolocal/shared/js/fn_numeros.js"></script> 
<script>
// ---------------------------------------------------------------------------------------
function validar(fecha)
{
	if (fecha.value == ""){
		alert('Ingrese una fecha');	
		return false
	}else{	
   	  if (validarfecha(fecha)){
	     return cantdias();
	  }else{
	     return false
	  }	
	}
}

// ---------------------------------------------------------------------------------------
function cantdias(){
	if (document.datos.vdiapeddesde.value!=''){
		if (document.datos.vdiapedhasta.value!=''){
			if (menorque(document.datos.vdiapeddesde.value,document.datos.vdiapedhasta.value) == false) {
				alert("La fecha desde no puede ser mayor que la fecha hasta.");
				document.datos.vdiapedcant.value="";
				return false;
				}
			else{
				//document.datos.vdiapedcant.value= contardias(document.datos.vdiapeddesde.value,document.datos.vdiapedhasta.value) +1;
				return true;
			}	
		 }
	}
}

function validarDesde(){
 if (document.datos.vdiapeddesde.value == ""){
	alert('Ingrese una fecha desde.');	
	return false;
 }else{		
    if (validarfecha(document.datos.vdiapeddesde)){
	   if (Trim(document.datos.vdiapedcant.value) != ''){
           if (validanumero(document.datos.vdiapedcant, 3, 0)){
		      if (parseInt(document.datos.vdiapedcant.value,10) < 1){
		         alert('La cantidad de dias debe ser mayor a cero.');
		         return false;			  
			  }else{
			     return true;
			  }
		   }else{
		      alert('La cantidad de dias debe ser numerica.');
		      return false;
		   }
	   }else{	   
	      return true;
	   }
	}else{
	   return false;
	}
 }	
}

// ---------------------------------------------------------------------------------------
function Validar_Formulario(){
	if (validarfecha(document.datos.vdiapeddesde) && validarfecha(document.datos.vdiapedhasta)){
	str = document.datos.vdiapeddesde.value;
	diad = str.substring(0, 2) ;
	mesd = str.substring(3, 5) ;
	anod = str.substring(6, 10) ;
	
	str = document.datos.vdiapedhasta.value;
	diah = str.substring(0, 2) ;
	mesh = str.substring(3, 5) ;
	anoh = str.substring(6, 10) ;

/*
	if (document.datos.tipvacnro.value == "") {
		alert("El período no tiene días correspondientes generados.");
		return false;
	}
*/	
	
	// VALIDAR FECHAS
	if (document.datos.vdiapeddesde.value == "") {
		alert("Debe ingresar la fecha Desde.");
		return false;
	}
	
	if (document.datos.vdiapedhasta.value == "") {
		alert("Debe ingresar la fecha Hasta.");
		return false;
	}
	if (anod > anoh) {
		alert("El año desde debe ser menor que el año hasta");
		return false;
	}
	if  ((mesd > mesh) && (anod == anoh)){
		alert("El mes desde debe ser menor que el mes hasta");	
		return false;
	}
	if  ((diad > diah) && (mesd == mesh)){
		alert("El dia desde debe ser menor que el dia hasta");
		return false;
	}
	if (isNaN(document.datos.vdiapedcant.value)) {
		alert("La cantidad de dias pedidos debe ser numerica.");
		return false;
	}
/*
	if (parseInt(document.datos.vdiapedcant.value) > parseInt(document.datos.diaspendientes.value)){
		alert("La cantidad de días pedidos no debe ser mayor a los días pendientes.");
		return false;
	}
*/

/*	
	if (!(menorque(document.datos.vacfecdesde.value, document.datos.vdiapeddesde.value) && menorque(document.datos.vdiapedhasta.value, document.datos.vacfechasta.value))){
		alert("Las fechas ingresadas no están incluidas en el período.");
		return false;
	}
*/	

	  Validar_Superposicion();
	}

}

function CalculaCorridos()
{
	str = document.datos.vdiapeddesde.value;
	diad = str.substring(0, 2) ;
	mesd = str.substring(3, 5) ;
	anod = str.substring(6, 10) ;
	str = document.datos.vdiapedhasta.value;
	diah = str.substring(0, 2) ;
	mesh = str.substring(3, 5) ;
	anoh = str.substring(6, 10) ;

  var fechadesde = new Date(anod,mesd,diad);
  var fechahasta = new Date(anoh,mesh,diah);
  var tiempo	 = fechahasta.getTime() - fechadesde.getTime();
  var dias = Math.floor(tiempo / (1000 * 60 * 60 * 24));
  if ((dias > 0) || (dias == 0))
    document.datos.diascorridos.value = dias + 1;
  else
    alert("Fecha Hasta menor que Fecha desde");
}

function Ayuda_Fecha(txt)
{
 var jsFecha = Nuevo_Dialogo(window, '/serviciolocal/shared/js/calendar.html', 16, 15);

 if (jsFecha == null) txt.value = ''
 else txt.value = jsFecha;
}


function Nuevo_Dialogo(w_in, pagina, ancho, alto)
{
 return w_in.showModalDialog(pagina,'', 'center:yes;dialogWidth:' + ancho.toString() + ';dialogHeight:' + alto.toString() + ';');
}

function calcularRango(tipo){
  if (validarDesde()){
     document.valida.location = 'pedido_emp_vac_calculo_gti_00.asp?tipo=' + tipo + '&tipovac=' + document.datos.tipvacnro.value + '&desde=' + document.datos.vdiapeddesde.value + '&hasta=' + document.datos.vdiapedhasta.value + '&cantidad=' + document.datos.vdiapedcant.value;
  }
}

function actualizarRango(hasta,cant, total,totalFer){
  document.datos.vdiapedcant.value  = cant;
  document.datos.vdiapedhasta.value = hasta;
  if (cant == ''){
      cant = 0;
  }
  document.datos.vdiaspedhabiles.value  = cant;
  document.datos.vdiaspednohabiles.value  = (parseInt(total) - parseInt(cant) - parseInt(totalFer));  
  document.datos.vdiaspedferiados.value = totalFer;
  document.datos.diascorridos.value = total;
}

function reCalcularRango(){
  if ((document.datos.vdiapedhasta.value != "") && (document.datos.vdiapedcant.value != "") &&  (document.datos.tipvacnro.value != "")){
     document.valida.location = 'pedido_emp_vac_calculo_gti_00.asp?tipo=SD&tipovac=' + document.datos.tipvacnro.value + '&desde=' + document.datos.vdiapeddesde.value + '&hasta=' + document.datos.vdiapedhasta.value + '&cantidad=' + document.datos.vdiapedcant.value + '&vdiapednro=<%= l_vdiapednro%>';
  }
}

function Validar_Superposicion(){
   document.valida.location = 'pedido_emp_vac_gti_06.asp?tipo=<%= l_tipo%>&desde=' + document.datos.vdiapeddesde.value + '&hasta=' + document.datos.vdiapedhasta.value + '&vdiapednro=' + document.datos.vdiapednro.value;
}

function rangoCorrecto(){
   document.datos.submit();
}

function rangoIncorrecto(){
   alert("El rango de fechas ingresado se superpone con otro pedido.");
}

function mostrarErrores(nroErr){
  switch (nroErr){
    case '1':{
	   alert('La cantidad de días es mayor a los disponibles');
	   break;
	}
  }
  document.datos.vdiapedcant.value  = '';
  document.datos.vdiapedhasta.value = '';
}

function actualizarTotales(cantidad,correspondientes){
  document.datos.totalcorr.value = correspondientes;
  document.datos.totalpedi.value = correspondientes - cantidad;
  document.datos.totalpend.value = cantidad;   
}

</script>

</head>
<body leftmargin="0" topmargin="0" rightmargin="0" bottommargin="0" onLoad="CargarFechas();">

<form name="datos" action="pedido_emp_vac_gti_03.asp?Tipo=<%=l_tipo%>" method="post">

<input type="hidden" name="tipo" value="<%=l_tipo%>">
<input type="hidden" name="vdiapednro" value="<%=l_vdiapednro%>">
<input type="hidden" name="tipvacnro" value="<%=l_tipvacnro%>" readonly>

<table border="0" cellpadding="0" cellspacing="0" height="100%">
<tr>
	<th colspan="5" align="left">Datos del Pedido</th>
	<th style="text-align: right;">
		<a class=sidebtnHLP href="Javascript:ayuda('<%= Request.ServerVariables("SCRIPT_NAME")%>');">Ayuda</a>
	</th>
</tr>

<tr valign="top">
   <td colspan="6" style="">
	&nbsp;<b>Empleado:</b>&nbsp;&nbsp;&nbsp;<%= l_terapenom %>
   </td>
</tr>

<tr>
  <td colspan="6" align="center">
    <table style="width:600px;border-color:gray ; border-width: 1 ; border-style:solid ; margin-top: 5px;">
	  <tr>
		<%' Buscar el periodos
		Set l_rs = Server.CreateObject("ADODB.RecordSet")
		l_sql = "SELECT vacacion.vacdesc, vacacion.vacnro, vacacion.vacfecdesde, vacacion.vacfechasta  "  
		l_sql = l_sql & " FROM  vacacion ORDER BY vacfecdesde DESC "
		rsOpen l_rs, cn, l_sql, 0 
		%>
		<td align="right"><b>Per&iacute;odo:</b></td>
		<td colspan="3">
		<select name=vacnro onChange="CargarFechas();reCalcularRango();">
			<%do while not l_rs.eof%>
			<option value="<%=l_rs("vacnro")%>"><%=l_rs("vacnro")%>&nbsp;-&nbsp;<%=l_rs("vacdesc")%></option>
			<%l_rs.MoveNext
			Loop
			l_rs.Close%>
		</select>		
		<%if l_tipo="M" then%>
			<script>
			  document.datos.vacnro.value = <%=l_vacnro%>
      	     <%if l_deshabilitar then%>
			 document.datos.vacnro.disabled = true;
			 <%end if%>
		    </script>			
		<%end if%>
		</td>
		<td align="right"><b>Desde:</b></td>
		<td ><input class="deshabinp" type="text" name="vacfecdesde" size="10" maxlength="10" readonly>
		</td>
		<td align="right"><b>Hasta:</b></td>
		<td ><input class="deshabinp" type="text" name="vacfechasta" size="10" maxlength="10" readonly>
		</td>
	  </tr>	
	  <tr>
  	    <td align="right"><b>D&iacute;as Corresp.:</b></td>
	    <td ><input class="deshabinp" type="text" name="vdiascorcant" size="4" maxlength="4"  readonly>
		</td>
		<td align="right" nowrap><b>D&iacute;as Ped.:</b></td>
		<td ><input class="deshabinp" type="text" name="diasyaped" size="4" maxlength="4" value="<%=l_diasyaped%>" readonly>
		</td>
		<td align="right" nowrap><b>D&iacute;as Pend.:</b></td>
		<td ><input class="deshabinp" type="text" name="diaspendientes" size="4" maxlength="4" value="<%'=l_repdesc%>" readonly>
		</td>
	    <td align="right" nowrap><b>Tip. Vac.:</b></td>
	    <td><input class="deshabinp" type="text" name="tipvacdesabr" size="10" maxlength="10" value="<%=l_tipvacdesabr%>" readonly>
	    </td>
	  </tr>

	</table>  
  </td>
</tr>
<tr>
  <td colspan="6" align="center">
    <table style="width:600px;border-color:gray ; border-width: 1 ; border-style:solid ; margin-top: 5px;">
		<tr>
   	    <%if l_deshabilitar then%>
		    <td align="center" colspan="8">
			    <b>Desde:</b>
				<input  type="text" name="vdiapeddesde" size="10" maxlength="10" value="<%= l_vdiapeddesde %>" readonly class="deshabinp">
				<a href="Javascript:;"><img src="/serviciolocal/shared/images/cal.gif" border="0"></a>	
                &nbsp;&nbsp;&nbsp;&nbsp;
			    <b>D&iacute;as Pedidos</b>
				<input type="text" name="vdiapedcant" size="4" maxlength="4" value="<%= l_vdiapedcant %>" readonly class="deshabinp">
                &nbsp;&nbsp;&nbsp;&nbsp;		        
				<b>Hasta:</b>
				<input type="text" name="vdiapedhasta" size="10" maxlength="10" value="<%= l_vdiapedhasta %>" readonly class="deshabinp">
				<a href="Javascript:;"><img src="/serviciolocal/shared/images/cal.gif" border="0"></a>	
			</td>		
		<%else%>
		    <td align="center" colspan="8">
			    <b>Desde:</b>
				<input  type="text" name="vdiapeddesde" size="10" maxlength="10" value="<%= l_vdiapeddesde %>" onChange="javascript:calcularRango('SD');">
				<a href="Javascript:Ayuda_Fecha(document.datos.vdiapeddesde);calcularRango('SD');"><img src="/serviciolocal/shared/images/cal.gif" border="0"></a>	
                &nbsp;&nbsp;&nbsp;&nbsp;
                <b>D&iacute;as Pedidos</b>
				<input type="text" name="vdiapedcant" size="4" maxlength="4" value="<%= l_vdiapedcant %>" onChange="javascript:calcularRango('SD');" >
                &nbsp;&nbsp;&nbsp;&nbsp;
                <b>Hasta:</b>
				<input type="text" name="vdiapedhasta" size="10" maxlength="10" value="<%= l_vdiapedhasta %>" onChange="if (validar(this)) calcularRango('CD');">
				<a href="Javascript:Ayuda_Fecha(document.datos.vdiapedhasta); calcularRango('CD');"><img src="/serviciolocal/shared/images/cal.gif" border="0"></a>	
			</td>
		<%end if%> 	
		</tr>
		<tr>
			<!--td align="right"><b>D&iacute;as Pedidos:</b></td>
			<td ><input type="text" name="vdiapedcant" size="4" maxlength="4" value="<%=l_vdiapedcant%>">
			</td-->
			<td align="right" nowrap><b>D&iacute;as Corrid.:</b></td>
			<td ><input class="deshabinp" type="text" name="diascorridos" size="4" maxlength="4" value="<%=l_diascorridos%>" readonly> 
			</td>
			<td align="right" nowrap><b>D&iacute;as H&aacute;b.:</b></td>
			<td ><input class="deshabinp" type="text" name="vdiaspedhabiles" size="4" maxlength="4" value="<%=l_vdiaspedhabiles%>" readonly>
			</td>
			<td align="right" nowrap><b>D&iacute;as No H&aacute;b.:</b></td>
			<td ><input class="deshabinp" type="text" name="vdiaspednohabiles" size="4" maxlength="4" value="<%=l_vdiaspednohabiles%>" readonly>
			</td>			
			<td align="right" nowrap><b>D&iacute;as Feri.:</b></td>
			<td ><input class="deshabinp" type="text" name="vdiaspedferiados" size="4" maxlength="4" value="<%=l_vdiaspedferiados%>" readonly>
			</td>
		</tr>		
	</table>
  </td>	
</tr>
<tr>
  <td colspan="6" align="center">
    <table style="width:600px;border-color:gray ; border-width: 1 ; border-style:solid ; margin-top: 5px;">
		<tr>
			<td align="center" colspan="8">
			<b>Total Corresp.:&nbsp;</b>
			<input class="deshabinp" type="text" name="totalcorr" size="4" maxlength="4" value="" readonly> 
			&nbsp;&nbsp;&nbsp;&nbsp;
			<b>Total Pedidos:&nbsp;</b>
			<input class="deshabinp" type="text" name="totalpedi" size="4" maxlength="4" value="" readonly> 
			&nbsp;&nbsp;&nbsp;&nbsp;
			<b>Total Pend.:&nbsp;</b>
			<input class="deshabinp" type="text" name="totalpend" size="4" maxlength="4" value="" readonly>
			</td>
		</tr>
	</table>
  </td>	
</tr>

<tr>
    <td align="right" class="th2" colspan="6">
		<%if l_deshabilitar then%>
		<a class=sidebtnABM href="Javascript:window.close()">Cerrar</a>		
		<%else%>
		<a class=sidebtnABM href=# onClick="Javascript:Validar_Formulario()">Aceptar</a>
		<a class=sidebtnABM href="Javascript:window.close()">Cancelar</a>
		<%end if%>
	</td>
</tr>
<iframe name="valida"  style="visibility=hidden;"  src="blanc.asp" width="500" height="500" ></iframe>
</table>

</form>


<script>
CargarFechas();
<% if l_tipo = "M" then%>
   reCalcularRango();
<%else%>
   calcularRango('SD');
<%end if%>
<%if l_deshabilitar then%>
   alert('El pedido tiene licencias generadas y no se puede modificar.');
<%end if%>
</script>
</body>
</html>

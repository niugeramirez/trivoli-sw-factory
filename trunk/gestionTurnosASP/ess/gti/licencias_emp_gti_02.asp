<% Option Explicit %>
<!--#include virtual="/serviciolocal/ess/shared/inc/const.inc"-->
<!--#include virtual="/serviciolocal/shared/db/conn_db.inc"-->
<!--#include virtual="/serviciolocal/shared/inc/fecha.inc"-->
<!--#include virtual="/serviciolocal/ess/shared/inc/accesoESS.inc"-->
<!--
-----------------------------------------------------------------------------
Archivo       : licencias_gti_02.asp
Descripcion   : Modulo que se encarga de mostrar los datos de las licencas
Creacion      : 24/03/2004
Autor         : Scarpa D.
Modificacion  :
  29/03/2004 - Scarpa D. - Correccion en el calculo del tope de licencias de vacaciones
  06/05/2004 - Scarpa D. - Se quitarin los campos de licencias parciales
  13/05/2004 - Scarpa D. - Se saco maternidad y lactancia
  18/10/2004 - Scarpa D. - Cambio en el calculo de los dias
  10-11-05 - Leticia A. - Si se configuro el ConfRep, mostrar los tipo de licencias configuradas.
  15/03/2006 - Mariano - Se quito la Vista V_EMPLEADO y se dejo la tabla
  30-07-2007 - Diego Rosso - Se agrego src="blanc.asp" para https.
-----------------------------------------------------------------------------
-->
<% 
on error goto 0

'Dim l_elfechacert
Dim l_emp_licnro
Dim l_tdnro
Dim l_tdnroant
Dim l_empleado
Dim l_elfechadesde
Dim l_elfechahasta
Dim l_elcantdias
Dim l_elcantdiashab
Dim l_elcantdiasfer
Dim l_eltipo
Dim l_elhoradesde
Dim l_elhorahasta
Dim l_elorden
Dim l_elmaxhoras
Dim l_tipo
Dim l_sql
Dim l_rs
Dim l_rs1
Dim l_vismednroenf
Dim l_licestnro
Dim l_estado

Dim l_ternro
Dim l_empleg
Dim l_apellido
Dim l_repnro 
Dim l_sql_confrep
Dim l_habilitar_estado
 
 ' ************
 l_repnro = 151 

Set l_rs  = Server.CreateObject("ADODB.RecordSet")

dim leg
leg = l_ess_empleg
l_ternro = l_ess_ternro

l_tipo = request("tipo")

'Si es el supervisor le permito modificar el estado
l_habilitar_estado = (Session("empleg") <> l_ess_empleg)

' ____________________________________________________________________
' Verificar si se cargaron Tipo de Licencias a mostrar en el ConfRep  
 l_sql = " SELECT repnro FROM confrep WHERE repnro=" & l_repnro
 rsOpen l_rs, cn, l_sql, 0 
 
 l_sql_confrep = ""
 if not l_rs.eof then  	' AND confrep.conftipo = 'TD' ?va
 	 l_sql_confrep = " INNER JOIN confrep ON confrep.confval=tipdia.tdnro  AND confrep.repnro="& l_repnro
 end if 
 l_rs.Close
 
 ' __________________________________________________________________

'===========================================================================================  	

l_sql = " SELECT * FROM empleado WHERE ternro=" & l_ternro
rsOpen l_rs, cn, l_sql, 0 

l_apellido = l_rs("terape") & " " & l_rs("terape2") & ", " & l_rs("ternom") & " " & l_rs("ternom2")
l_empleg   = leg			

l_rs.close

select Case l_tipo
	Case "A":
		l_emp_licnro = ""
		l_tdnro = ""
		l_elfechadesde = ""
		l_elfechahasta = ""
		l_elcantdias = ""
		l_eltipo = 1
		'l_elfechacert = ""
		l_licestnro = 1

	Case "M", "C":  'C es consulta.
		l_emp_licnro = request("cabnro")
		l_sql = "SELECT emp_licnro, tdnro, empleado, elfechadesde, elfechahasta,licestnro "
		'l_sql = l_sql & " , elcantdias, elmaxhoras, elorden, eltipo, elhoradesde, elhorahasta, empleg, elfechacert "
		l_sql = l_sql & " , elcantdias, empleg "
	    l_sql = l_sql & " FROM emp_lic "
		l_sql = l_sql & " INNER JOIN empleado ON empleado.ternro = emp_lic.empleado "
		l_sql  = l_sql  & "WHERE emp_licnro = " & l_emp_licnro
		rsOpen l_rs, cn, l_sql, 0 
		if not l_rs.eof then
			l_emp_licnro = l_rs("emp_licnro")
			l_tdnro    = l_rs("tdnro")
			l_tdnroant = l_rs("tdnro") ' hay que gusradr el anterior por si lo cambia....
			l_empleado = l_rs("empleg")
			l_elfechadesde = l_rs("elfechadesde")
			l_elfechahasta = l_rs("elfechahasta")
			l_elcantdias = l_rs("elcantdias")
			if isnull(l_rs("licestnro")) then
			   l_licestnro = 1
			else
			   l_licestnro = l_rs("licestnro")
			end if
			'l_elfechacert = l_rs("elfechacert")
			' para complementos
		end if
		l_rs.Close
end select

%>

<html>
<head>
<link href="../<%= c_estilo %>" rel="StyleSheet" type="text/css">
<meta http-equiv="Content-Type" content="text/html; charset=iso-8859-1">
<title>Licencias - Gesti&oacute;n de Tiempos - RHPro &reg;</title>
</head>
<script src="/serviciolocal/shared/js/fn_windows.js"></script>
<script src="/serviciolocal/shared/js/fn_ayuda.js"></script>
<script src="/serviciolocal/shared/js/fn_ay_generica.js"></script>
<script src="/serviciolocal/shared/js/fn_hora.js"></script>
<script src="/serviciolocal/shared/js/fn_fechas.js"></script>
<script src="/serviciolocal/shared/js/fn_buscar_emp.js"></script>
<script src="/serviciolocal/shared/js/fn_help_emp.js"></script>

<%
Dim l_tipAutorizacion  'Es el tipo del circuito de firmas
Dim l_HayAutorizacion  'Es para ver si las autorizaciones estan activas
Dim l_PuedeVer         'Es para ver si las autorizaciones estan activas

l_tipAutorizacion = 6  'Es del tipo licencias

l_sql = "select * from cystipo "
l_sql = l_sql & "where (cystipo.cystipact = -1) and cystipo.cystipnro = " & l_tipAutorizacion 

rsOpen l_rs, cn, l_sql, 0 

l_HayAutorizacion = not l_rs.eof

' BORRARRRRRRRRRRRRRRRRRRRRRRRRRRRRRRRRRRRRRRRRRRRRRRRRR!!!!!!!!!!!!!!
'l_HayAutorizacion = false

l_rs.close

if l_HayAutorizacion AND (l_tipo = "M") then

  l_sql = "select cysfirautoriza, cysfirsecuencia, cysfirdestino from cysfirmas "
  l_sql = l_sql & "where cysfirmas.cystipnro = " & l_tipAutorizacion & " and cysfirmas.cysfircodext = '" & l_emp_licnro & "' " 
  l_sql = l_sql & "order by cysfirsecuencia desc"

  rsOpen l_rs, cn, l_sql, 0 

  l_PuedeVer = false

  if not l_rs.eof then
    if (l_rs("cysfirautoriza") = session("UserName")) or (l_rs("cysfirdestino") = session("UserName")) then 
	   'Es una modificación del ultimo o es el nuevo que autoriza 
       l_PuedeVer = True 
    end if
  end if
  l_rs.close
  If not l_PuedeVer then
    response.write "<script>alert('No esta autorizado a ver o modificar este registro.');window.close()</script>"
	response.end
  End if
End if
%>
<script>

function cargarIfrm(){
	document.ifrmcomp.datos.emp_licnro.value   = document.datos.emp_licnro.value;
	document.ifrmcomp.datos.tdnro.value        = document.datos.tdnro.value;
	document.ifrmcomp.datos.tdnroant.value     = document.datos.tdnroant.value;
	document.ifrmcomp.datos.empleado.value     = document.datos.ternro.value;
	document.ifrmcomp.datos.elfechadesde.value = document.datos.elfechadesde.value;
	document.ifrmcomp.datos.elfechahasta.value = document.datos.elfechahasta.value;
	document.ifrmcomp.datos.elcantdias.value   = document.datos.elcantdias.value;
	document.ifrmcomp.datos.seleccion.value    = document.datos.seleccion.value;
	document.ifrmcomp.datos.seleccion1.value   = document.datos.seleccion1.value;
<%if l_habilitar_estado then%>
	document.ifrmcomp.datos.licestnro.value   = document.datos.licestnro.value;
<%end if%>
	
}

window.resizable = 'no';

// RUTINAS DE VALIDACION ==================================================================
function menorque(fecha1,fecha2){
	var f1= new Date(); 
	f1.setFullYear(fecha1.substr(6,4),fecha1.substr(3,2)-1,fecha1.substr(0,2));
	var segf1=Date.parse(f1); 

	var f2= new Date(); 
	f2.setFullYear(fecha2.substr(6,4),fecha2.substr(3,2)-1,fecha2.substr(0,2));
	var segf2=Date.parse(f2); 

	if ((segf1<segf2)||(fecha1==fecha2)){return true}
	else{return false}
}

function validar(fecha){
	if (fecha.value == "")
		alert('Ingrese una fecha');	
	if (validarfecha(fecha)) 
		{cantdias()};	
}

function Validar_Formulario(){
var errores = 0;
var guardarDatos = 1;

<% if l_HayAutorizacion then ' Si se debe tomar autorizacion %>
// Verifico que se haya cargado la autorización 
if ((errores == 0) && (((document.datos.seleccion.value == "") && (document.datos.seleccion1.value == "")) && ("<%= l_tipo %>" == "A"))){
    alert("Debe ingresar una autorización.");
	errores++;
}
<% End If %>

if ((errores == 0) && (document.datos.tdnro.value == "")){
	alert("Debe ingresar un tipo de Licencia.");
	errores++;
}

if ((errores == 0) && (document.datos.elfechadesde.value == "")){
	alert("Debe ingresar la fecha desde.");
	errores++;
}

if ((errores == 0) && (document.datos.elfechahasta.value == "")){ 
	alert("Debe ingresar la fecha hasta.");
	errores++;
}

if ((errores == 0) && !(validarfecha(document.datos.elfechadesde) && validarfecha(document.datos.elfechahasta))){
	errores++;
}

if ((errores == 0) && (menorque(document.datos.elfechadesde.value,document.datos.elfechahasta.value) == false)){ 
	alert("la fecha desde no puede ser mayor que la fecha hasta.");
	errores++;
}

if ((errores == 0) && (document.datos.elcantdias.value == "")){
	alert("Debe ingresar la cantidad de días.");
	errores++;
}

<%if l_habilitar_estado then%>
if ((errores == 0) && (document.datos.licestnro.value == "")){
	alert("Debe selecionar un estado.");
	errores++;
}
<%end if%>

if ((errores == 0) && (!document.ifrmcomp.ValidarDatos())){
  //Imprime un error indicando que los datos del complemento son incorrectos
	errores++;
}

if (errores == 0){
//   document.valida.location = 'licencias_emp_gti_05.asp?desde=' + document.datos.elfechadesde.value + '&hasta=' + document.datos.elfechahasta.value + '&tipo=<%= l_tipo%>&emplicnro=<%= l_emp_licnro%>&ternro=' + document.datos.ternro.value + '&tdnro=' + document.datos.tdnro.value + '&cantidad=' + document.datos.elcantdias.value;   
   if (document.ifrmcomp.document.getElementById("hayparams")){
      document.valida.location = 'licencias_emp_gti_05.asp?empleg=<%= request.querystring("empleg") %>&desde=' + document.datos.elfechadesde.value + '&hasta=' + document.datos.elfechahasta.value + '&tipo=<%= l_tipo%>&emplicnro=<%= l_emp_licnro%>&ternro=' + document.datos.ternro.value + '&tdnro=' + document.datos.tdnro.value + '&cantidad=' + document.datos.elcantdias.value + document.ifrmcomp.params();      
   }else{
      document.valida.location = 'licencias_emp_gti_05.asp?empleg=<%= request.querystring("empleg")%>&desde=' + document.datos.elfechadesde.value + '&hasta=' + document.datos.elfechahasta.value + '&tipo=<%= l_tipo%>&emplicnro=<%= l_emp_licnro%>&ternro=' + document.datos.ternro.value + '&tdnro=' + document.datos.tdnro.value + '&cantidad=' + document.datos.elcantdias.value;   
   }   
}

}
// RUTINAS AUXILIARES ======================================================================

function guardarLicencia(){

  cargarIfrm();
  abrirVentanaH('','vent_oculta',200,200);	
  document.ifrmcomp.datos.submit();
  //document.datos.submit();

}

function cantdias(){
	if (document.datos.elfechadesde.value!=''){
		if (document.datos.elfechahasta.value!=''){
			if (menorque(document.datos.elfechadesde.value,document.datos.elfechahasta.value) == false) {
				alert("la fecha desde no puede ser mayor que la fecha hasta.");
				document.datos.elcantdias.value="";
				}
			else{
			   if (document.ifrmcomp.document.getElementById("hayparams")){
			      document.valida.location = 'licencias_emp_gti_06.asp?empleg=<%= request.querystring("empleg") %>&desde=' + document.datos.elfechadesde.value + '&hasta=' + document.datos.elfechahasta.value + '&tipo=DIFERENCIA' + '&tdnro=' + document.datos.tdnro.value + '&ternro=' + document.datos.ternro.value + document.ifrmcomp.params();      
			   }else{
			      document.valida.location = 'licencias_emp_gti_06.asp?empleg=<%= request.querystring("empleg") %>&desde=' + document.datos.elfechadesde.value + '&hasta=' + document.datos.elfechahasta.value + '&tipo=DIFERENCIA' + '&tdnro=' + document.datos.tdnro.value + '&ternro=' + document.datos.ternro.value;   						
			   }
//			  mostrarComplemento();
			}
		 }
	}
}


function aumentardias(){
if (isNaN(document.datos.elcantdias.value) || (document.datos.elcantdias.value<1)){
		alert('La cantidad de días no es correcta');
		document.datos.elcantdias.value="";
	}	
else {
	fechadesde=document.datos.elfechadesde.value;
  	if ((document.datos.elcantdias.value=="")||(isNaN(document.datos.elcantdias.value)))
		document.datos.elcantdias.value=1;
	else{
	   if (document.ifrmcomp.document.getElementById("hayparams")){
           document.valida.location = 'licencias_emp_gti_06.asp?empleg=<%= request.querystring("empleg") %>&desde=' + document.datos.elfechadesde.value + '&cant=' + document.datos.elcantdias.value + '&tipo=SUMAR' + '&tdnro=' + document.datos.tdnro.value + '&ternro=' + document.datos.ternro.value + document.ifrmcomp.params();      
	   }else{
           document.valida.location = 'licencias_emp_gti_06.asp?empleg=<%= request.querystring("empleg") %>&desde=' + document.datos.elfechadesde.value + '&cant=' + document.datos.elcantdias.value + '&tipo=SUMAR' + '&tdnro=' + document.datos.tdnro.value;   	
	   }

//	   mostrarComplemento();
    } 
}
}

// RUTINAS DE FECHAS ======================================================================
function Nuevo_Dialogo(w_in, pagina, ancho, alto)
{
 return w_in.showModalDialog(pagina,'', 'center:yes;dialogWidth:' + ancho.toString() + ';dialogHeight:' + alto.toString() + ';');
}
function Ayuda_Fecha(txt)
{
 var jsFecha = Nuevo_Dialogo(window, '/serviciolocal/shared/js/calendar.html', 16, 15);
 if (jsFecha == null) txt.value = ''
 else txt.value = jsFecha;
}
// Para llamar a control de firmas, mandandole la descripcion y demas ======================
function Firmas()  
{
  abrirVentana('cysfirmas_00.asp?obj=document.datos.seleccion&amp;tipo=<%= l_tipAutorizacion %>&amp;codigo=<%= l_emp_licnro %>&amp;descripcion=' + datos.tdnro(datos.tdnro.selectedIndex).text ,'_blank','toolbar=no,location=no,directories=no,status=no,menubar=no,scrollbars=no,resizable=yes,width=421,height=180')
}

// MOSTRAR empleado ========================================================================
function Mostrar(){
verificacodigo(document.datos.empleado,document.datos.descripcion,'empleg','terape, ternom','empleado');
return false;
}

function ActualizarSelect(){
//if ((document.datos.descripcion.value!==""))
//	if ((document.datos.tdnro.value==8) && (document.datos.vismednroenf.value=="" ||document.datos.vismednroenf.value=="null"))
//		CargarVisitas(document.datos.vismednroenf,"<%=l_vismednroenf%>");
}

function Chequearempleado() {
	if (document.datos.empleado.value==""){
		alert('Ingrese un empleado.');
		return 0;
	}else{
	    return 1;
	}
}

function mostrarComplemento(){
  var desde,hasta;
  
  desde = document.datos.elfechadesde.value;
  hasta = document.datos.elfechahasta.value;

  if (Chequearempleado()){
     //alert('comp_lic_emp_carga_gti_00.asp?tipo=<%= l_tipo%>&tdnro=' + document.datos.tdnro.value + '&emp_licnro=' + document.datos.emp_licnro.value + '&ternro=' + document.datos.ternro.value + '&empleg=' + document.datos.empleado.value);
//     if (document.ifrmcomp.document.getElementById("hayparams")){
//        document.ifrmcomp.recargar('<%= l_tipo%>',document.datos.tdnro.value,document.datos.emp_licnro.value,document.datos.ternro.value,document.datos.empleado.value,desde,hasta);
//	 }else{
//        document.ifrmcomp.location = 'comp_lic_emp_carga_gti_00.asp?tipo=<%= l_tipo%>&tdnro=' + document.datos.tdnro.value + '&emp_licnro=' + document.datos.emp_licnro.value + '&ternro=' + document.datos.ternro.value + '&empleg=' + document.datos.empleado.value + '&desde=' + desde + '&hasta=' + hasta;
	// }
	 
	 if (document.datos.tdnro.value == '2'){
	    document.ifrmcomp.location = 'comp_lic_emp_vacacion_gti_00.asp?tipo=<%= l_tipo%>&tdnro=' + document.datos.tdnro.value + '&emp_licnro=' + document.datos.emp_licnro.value + '&ternro=' + document.datos.ternro.value + '&empleg=<%= request.querystring("empleg") %>&desde=' + desde + '&hasta=' + hasta;	 
	 }else{
	    document.ifrmcomp.location = 'comp_lic_emp_carga_gti_00.asp?tipo=<%= l_tipo%>&tdnro=' + document.datos.tdnro.value + '&emp_licnro=' + document.datos.emp_licnro.value + '&ternro=' + document.datos.ternro.value + '&empleg=<%= request.querystring("empleg")%>&desde=' + desde + '&hasta=' + hasta;
	 }
  }
}

//function cambiarEstadoCert(){
//  if (document.datos.elfechacertcheck.checked){
//	  document.datos.elfechacert.disabled  = 0;
//	  document.datos.elfechacert.className = 'habinp';
//  }else{
//	  document.datos.elfechacert.disabled  = 1;
//	  document.datos.elfechacert.className = 'deshabinp';  
//  }
//}

function cambioIframe(){
	if (isNaN(document.datos.elcantdias.value) || (document.datos.elcantdias.value<1)){
		document.datos.elcantdias.value="";
	}else {
		fechadesde=document.datos.elfechadesde.value;
	  	if ((document.datos.elcantdias.value=="")||(isNaN(document.datos.elcantdias.value)))
			document.datos.elcantdias.value=1;
		else{
		   if (document.ifrmcomp.document.getElementById("hayparams")){
	           document.valida.location = 'licencias_gti_11.asp?empleg=<%= request.querystring("empleg") %>&desde=' + document.datos.elfechadesde.value + '&cant=' + document.datos.elcantdias.value + '&tipo=SUMAR' + '&tdnro=' + document.datos.tdnro.value + '&ternro=' + document.datos.ternro.value + document.ifrmcomp.params();      
		   }else{
	           document.valida.location = 'licencias_gti_11.asp?empleg=<%= request.querystring("empleg") %>&desde=' + document.datos.elfechadesde.value + '&cant=' + document.datos.elcantdias.value + '&tipo=SUMAR' + '&tdnro=' + document.datos.tdnro.value;   	
		   }
		}	
	}
}

</script>

<body leftmargin="0" rightmargin="0" topmargin="0" bottommargin="0">  

<form name="datos" action="licencias_emp_gti_03.asp?empleg=<%= request.querystring("empleg") %>&tipo=<%= l_tipo %>" method="post">

<input type="Hidden" name="emp_licnro"	value="<%= l_emp_licnro %>">
<input type="Hidden" name="tdnroant"	value="<%= l_tdnroant %>">
<input type="Hidden" name="ternro"     value="<%= l_ternro %>">
<input type="Hidden" name="emplegant" value="<%= l_empleg %>">

<table cellspacing="0" cellpadding="0" border="0" width="100%" height="100%">
  <tr>
    <th class="th2">Datos de Licencia</td>
    <th class="th2" align="right"  colspan="5">&nbsp;</th>
  </tr>
<tr>
    <td align="right"><b>Empleado:</b></td>
	<td colspan="5">
	<input type="text" name="empleado" readonly  class="deshabinp" size="10" maxlength="10" value="<%=l_empleg%>">
	&nbsp;&nbsp;
	<input type="Text" name="descripcion" size="40" readonly class="deshabinp" value="<%= l_apellido%>">
	</td>
</tr>

<tr>
    <td align="right"><b>Tipo de Licencia:</b></td>
	<td colspan="5">
	<select name="tdnro" size="1" onchange="JavaScript:mostrarComplemento();">
      <option value="">&laquo;Seleccione una opci&oacute;n&raquo;</option>	
	  <%Set l_rs = Server.CreateObject("ADODB.RecordSet")
	    l_sql = "SELECT tdnro, tddesc "
		l_sql  = l_sql  & " FROM tipdia "
		'La descripcion de los tipos de dias son(en forma correlativa):
		'matrimonio, estudio, mudanza, vacaiones, semana turismo, interv. quirurgica, exam. ginecologico
		'donacion sangre, matrimonio familiares, duelo, interv. quirurgica familiar
		'l_sql  = l_sql  & " WHERE tdnro IN (4,7,19,2,18,35,25,23,31,5,34)"
		if l_sql_confrep <> "" then
			l_sql = l_sql & l_sql_confrep
		end if
		l_sql  = l_sql  & " ORDER BY tddesc "
		rsOpen l_rs, cn, l_sql, 0 
		do until l_rs.eof	 %>	
		<option value="<%= l_rs("tdnro")%>"><%=l_rs("tddesc") %> </option>
		<%l_rs.Movenext
		loop
		l_rs.Close
		set l_rs= nothing 		%>	
	</select>
	<script>
		document.datos.tdnro.value = "<%= l_tdnro %>";
	</script>
	</td>
</tr>
<tr>
    <td align="right"><b>Desde:</b></td>
	<td>
	<input  type="text" name="elfechadesde" size="10" maxlength="10" value="<%= l_elfechadesde %>" onblur="validar(this);mostrarComplemento();">
	<a href="Javascript:Ayuda_Fecha(document.datos.elfechadesde); cantdias();"><img src="/serviciolocal/shared/images/cal.gif" border="0"></a>	
	</td>
	<td align="right"><b>Cant. de D&iacute;as</b></td>
	<td>
	<input type="text" name="elcantdias" size="4" maxlength="4" value="<%= l_elcantdias %>" onblur="javascript:aumentardias();" >
	</td>
    <td align="right"><b>Hasta:</b></td>
	<td>
	<input type="text" name="elfechahasta" size="10" maxlength="10" value="<%= l_elfechahasta %>" onblur="validar(this);">
	<a href="Javascript:Ayuda_Fecha(document.datos.elfechahasta); cantdias(); "><img src="/serviciolocal/shared/images/cal.gif" border="0"></a>	
	</td>
</tr>
<%if l_habilitar_estado then%>
<tr>
   <td align="right">
     <b>Estado:</b>
   </td>
   <td>
	<select name="licestnro" size="1">
	<option value="">&laquo;Seleccione una Opci&oacute;n&raquo;</option>	
		<%
		Set l_rs = Server.CreateObject("ADODB.RecordSet")
	    l_sql = "SELECT * FROM lic_estado ORDER BY licestdesabr "
		rsOpen l_rs, cn, l_sql, 0 
		
		do until l_rs.eof
			l_estado = l_rs("licestnro")
			if isNull(l_estado) then
		      l_estado = ""
			end if
			%><option value="<%= l_estado%>" <%if CStr(l_estado) = CStr(l_licestnro) then response.write "selected" end if %>><%=l_rs("licestdesabr") & " (" & l_estado & ")" %> </option><%
			l_rs.Movenext
		loop
		l_rs.Close
		set l_rs= nothing
		%>	
	</select>
     
   </td>
</tr>
<%end if%>
<tr>
<td colspan="6" height="50%"><b>Datos Complementarios:</b>
  <iframe name="ifrmcomp" src="comp_lic_emp_carga_gti_00.asp?empleg=<%= request.querystring("empleg") %>" width="100%" height="100%" scrolling="No" frameborder="0"> </iframe>
</td>
<tr>
    <td align="right" class="th2" colspan="6">
	    <input type="hidden" name="seleccion" value="">
	    <input type="hidden" name="seleccion1" value="">
<% if l_HayAutorizacion then ' Si se debe tomar autorizacion %>
		<a class=sidebtnSHW href="Javascript:Firmas();">Autorizar</a>
<% End If %>		
		<a class=sidebtnABM href="Javascript:Validar_Formulario()">Aceptar</a>
		<a class=sidebtnABM href="Javascript:window.close()">Cancelar</a>
	</td>
</tr>
</table>

</form>

<iframe name="valida" src="blanc.asp" width="0" height="0" ></iframe>

<script> 
<%if l_tipo = "M" then%>
   mostrarComplemento();
<%end if%>
</script>

</body>
</html>

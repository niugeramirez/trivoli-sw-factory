<% Option Explicit %>
<!--#include virtual="/ess/ess/shared/inc/const.inc"-->
<!--#include virtual="/serviciolocal/shared/inc/const_eva.inc"-->
<html>
<head>
<link href="../<%=c_estilo %>" rel="StyleSheet" type="text/css">
<meta http-equiv="Content-Type" content="text/html; charset=iso-8859-1">
<title>Reportes - Gesti&oacute;n de Desempe&ntilde;o - RHPro &reg;</title>
</head>
<script src="/serviciolocal/shared/js/fn_windows.js"></script>
<script>



function buscemp(){
var evaevenro;
var evaevedesabr='';
	if (parent.document.datos.evaevenro.value=="0")
		alert('Debe seleccionar un evento.');
	else{
		evaevenro=parent.document.datos.evaevenro.value;
		JavaScript:abrirVentana('form_carga_eva_06.asp?evaevenro='+evaevenro +'&evaevedesabr='+evaevedesabr,'',450,400);
	}
}

function buscartab(esto){
var evaevenro;
evaevenro=parent.document.datos.evaevenro.value;
if (parent.document.datos.evaevenro.value=="0")
	alert('Debe seleccionar un evento.');
else
{
	if (isNaN(esto)){
		esto = "";
		<%if ccodelco=-1 then%>
		alert("El nùmero ingresado no es correcto.");
		<%else%>
		alert("El legajo ingresado no es correcto.");		
		<%end if%>
	}
	else 
		abrirVentanaH('nuevo_emp.asp?empleg='+esto+'&tabla=empleado INNER JOIN evacab ON (empleado.ternro=evacab.empleado and evacab.evaevenro='+evaevenro+')','',200,100);
}	
}

function Tecla(num){
  var evaevenro;
  evaevenro=parent.document.datos.evaevenro.value;
  if (num==13) {
	if (parent.document.datos.evaevenro.value=="0")
		alert('Debe seleccionar un evento.');
	else
		abrirVentanaH('nuevo_emp.asp?empleg='+document.datos.empleg.value+'&tabla=empleado INNER JOIN evacab ON (empleado.ternro=evacab.empleado and evacab.evaevenro='+evaevenro+')','',200,100);
	return false;
  }
  return num;
}

function nuevoempleado(ternro,empleg,terape,ternom)
{
if (empleg != 0) {	
			document.datos.ternro.value = ternro;
			document.datos.empleg.value = empleg;
			document.datos.empleado.value = terape + ", " + ternom;
}
else
{
	<%if ccodelco=-1 then%>
	alert('Supervisado incorrecto');
	<%else%>
	alert('Empleado	incorrecto');	
	<%end if%>
}
}

</script>

<body leftmargin="0" rightmargin="0" topmargin="0" bottommargin="0">
<form name="datos"  action="#" method="post">
<input type="Hidden" name="ternro" value="">
<table cellspacing="0" cellpadding="0" border="0" width="100%">
<tr>
	<td colspan="2" height="18"></td>
</tr>
<tr>
   <td align="right">
   		<b><font size=-2><%if ccodelco=-1 then%>Supervisado<%else%>Empleado<%end if%>:</b>
	</td>	
	<td>
		<input type="text" size="8" name="empleg" onKeyPress="return Tecla(event.keyCode)" onchange="buscartab(this.value);"> 
			<a onclick="JavaScript:buscemp()" onmouseover="window.status='Buscar por Apellido'" onmouseout="window.status=' '" style="cursor:hand;">
				<img align="absmiddle" src="/serviciolocal/shared/images/profile.gif" alt="Ayuda Empleados" border="0">
			</a>
		<br>
		<input class="rev" style="background : #e0e0de;" readonly type="text" name="empleado" size="40" maxlength="35" value="">
	</td>
</tr>
<tr>
	<td colspan="2" height="35"></td>
</tr>

</table>
</form>
</body>
</html>

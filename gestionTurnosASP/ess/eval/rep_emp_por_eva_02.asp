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
	window.open('help_emp_01.asp','new','toolbar=no,location=no,directories=no,satus=no,menubar=no,scrollbars=no,resizable=no,width=700,height=400');
}

function buscartab(esto){
if (isNaN(esto)){
	esto = "";
	<%if ccodelco=-1 then%>
	alert("El nùmero ingresado no es correcto.");
	<%else%>
	alert("El legajo ingresado no es correcto.");
	<%end if%>
	}
else {
	abrirVentanaH('nuevo_emp.asp?empleg='+esto+'&tabla=empleado INNER JOIN evadetevldor ON evadetevldor.evaluador=empleado.ternro','',200,100);
	}
}

function Tecla(num){
  if (num==13) {
  		buscartab(document.datos.empleg.value);
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
	alert('Supervisor incorrecto');
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
   		<b><%if ccodelco=-1 then%>Supervisor<%else%>Evaluador<%end if%>:</b>
	</td>	
	<td>
	<input type="text" size="8" name="empleg" onKeyPress="return Tecla(event.keyCode);"  onchange="buscartab(this.value);"> 
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
<%'onchange="javascript:abrirVentanaH('nuevo_emp.asp?empleg='+document.datos.empleg.value,'',200,100);"%>
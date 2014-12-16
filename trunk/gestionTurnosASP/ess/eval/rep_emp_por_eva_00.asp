<% Option Explicit %>
<!--#include virtual="/ess/ess/shared/inc/const.inc"-->
<%
'Modificado: 17-08-2005 CCRossi Agrandar llamado a ventana filtro estructuras
%>
<html>
<head>
<link href="../<%=c_estilo %>" rel="StyleSheet" type="text/css">
<meta http-equiv="Content-Type" content="text/html; charset=iso-8859-1">
<title>Reportes - Gesti&oacute;n de Desempe&ntilde;o - RHPro &reg;</title>
</head>
<script src="/serviciolocal/shared/js/fn_windows.js"></script>
<script>
function ponerestructuras(niv1,tex1,niv2,tex2,niv3,tex3){
	document.datos.estrnro1.value=niv1;
	document.datos.nivel1.value=tex1;

	document.datos.estrnro2.value=niv2;
	document.datos.nivel2.value=tex2;

	document.datos.estrnro3.value=niv3;
	document.datos.nivel3.value=tex3;
}
function deshabil_niv_estr(){
	document.all.filtro_estr.className="sidebtnDSB"
	document.datos.nivel1.disabled= true;
	document.datos.nivel1.className="deshabinp";
	document.datos.nivel2.disabled= true;
	document.datos.nivel2.className="deshabinp";
	document.datos.nivel3.disabled= true;
	document.datos.nivel3.className="deshabinp";
}
function habil_niv_estr(){
		document.all.filtro_estr.className="sidebtnSHW"
		document.datos.nivel1.disabled= false;
		document.datos.nivel1.className="habinp";
		document.datos.nivel2.disabled= false;
		document.datos.nivel2.className="habinp";
		document.datos.nivel3.disabled= false;
		document.datos.nivel3.className="habinp";
}
function clickcheck(){

	if (document.datos.check.checked == true){
		deshabil_niv_estr()
	}
	else{
		habil_niv_estr()
	}
}

function filtrarestr(){
	if (document.all.filtro_estr.className=="sidebtnSHW")
		abrirVentana('rep_filtro_estr_eva_00.asp','',550,220);
}
</script>

<body leftmargin="0" rightmargin="0" topmargin="0" bottommargin="0">
<form name="datos"  action="#" method="post">
<table cellspacing="0" cellpadding="0" border="0" width="100%">
	<tr>
		<td align="center" colspan="4">
		<input type="checkbox" name="check" checked onclick="javascript:clickcheck();">
		<b>Todas las Estructuras</b>
		&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;
		<a name="filtro_estr" class=sidebtnDSB href="Javascript:filtrarestr();">Filtrar</a>
		</td>
					
	</tr>
	<tr>
		    <td align="right"><b>Nivel 1:</b></td>
			<td colspan="3">
				<input type="Hidden" name="estrnro1" value="">
				<input readonly class="deshabinp" type="text" name="nivel1" size="40" maxlength="45" value="">
			</td>
	</tr>
	<tr>
		    <td align="right"><b>Nivel 2:</b></td>
			<td colspan="3">
				<input type="Hidden" name="estrnro2" value="">
				<input readonly class="deshabinp" type="text" name="nivel2" size="40" maxlength="45" value="">
			</td>
	</tr>
	<tr>
		    <td align="right"><b>Nivel 3:</b></td>
			<td colspan="3">
				<input type="Hidden" name="estrnro3" value="">
				<input readonly class="deshabinp" type="text" name="nivel3" size="40" maxlength="45" value="">
			</td>
	</tr>
</table>
</form>
</body>
</html>

<% Option Explicit %>
<!--#include virtual="/serviciolocal/shared/db/conn_db.inc"-->
<!--#include virtual="/serviciolocal/shared/inc/sec.inc"-->
<!--#include virtual="/serviciolocal/shared/inc/const.inc"-->
<!--#include virtual="/serviciolocal/shared/inc/fecha.inc"-->

<% on error goto 0
Dim l_rs
Dim l_sql

Dim l_fecini
Dim l_fecfin

Dim l_buqdes

l_buqdes 	  = request.querystring("qbuqdes")

%>


<html>
<head>
<link href="/serviciolocal/ess/shared/css/tables_gray.css" rel="StyleSheet" type="text/css">
<title><%= Session("Titulo")%> Estad�sticas </title>
<script src="/serviciolocal/shared/js/fn_windows.js"></script>
<script src="/serviciolocal/shared/js/fn_confirm.js"></script>
<script src="/serviciolocal/shared/js/fn_ayuda.js"></script>
<script src="/serviciolocal/shared/js/fn_fechas.js"></script>
<script src="/serviciolocal/shared/js/fn_ay_generica.js"></script>
<script>
function Imprimir(){
	parent.frames.ifrm.focus();
	window.print();	
}

function Actualizar(destino){

	document.ifrm.location = "rep_rem_detalle_buques_04.asp?qbuqdes=<%= l_buqdes %>" ;			
	
}

function Ayuda_Fecha(txt){
	var jsFecha = Nuevo_Dialogo(window, '/serviciolocal/shared/js/calendar.html', 16, 15);
	if (jsFecha == null){
		//txt.value = '';
	}else{
		txt.value = jsFecha;
		//DiadeSemana(jsFecha);
	}
}



</script>


</head>
<body leftmargin="0" topmargin="0" rightmargin="0" bottommargin="0">
<form name="datos">
<table border="0" cellpadding="0" cellspacing="0" height="100%">
	<tr style="border-color :CadetBlue;">
		<td align="left" class="barra" nowrap>
			<!--<a class=sidebtnSHW href="Javascript:window.close();">Salir</a>--></td>
		<td align="right" class="barra" colspan="5">
			<a class=sidebtnSHW href="Javascript:Actualizar('ifrm')">Actualizar</a>		  
			<a class=sidebtnSHW href="Javascript:Imprimir()">Imprimir</a>		  
			<!--<a class=sidebtnSHW href="Javascript:Actualizar('exel')">Excel</a> -->
			&nbsp;
			
		</td>
	</tr>
	<!--
	<tr>
		<td align="right" size="10%">
			<b>Fecha Desde:</b>
		</td>
		<td>
			<input  type="text" name="fecini" size="10" maxlength="10" value="<%'= l_fecini %>" >
			<a href="Javascript:Ayuda_Fecha(document.datos.fecini);"><img src="/serviciolocal/shared/images/cal.gif" border="0"></a>
		</td>
		<td align="right" size="10%">
			<b>Fecha Hasta:</b>
		</td>	
		<td>
			<input  type="text" name="fecfin" size="10" maxlength="10" value="<%'= l_fecfin %>">
			<a href="Javascript:Ayuda_Fecha(document.datos.fecfin);"><img src="/serviciolocal/shared/images/cal.gif" border="0"></a>
		</td>

		<td  align="right" nowrap><b>&nbsp;</b></td>
		<td><select name="repnro" size="1" style="width:30;">
				<option value=0 selected>Todos</option>
				<option value=1>1-Exportaci�n - Detalle de Buques</option>
				<option value=2>2-Importaci�n - Detalle de Buques </option>
				<option value=3>3-Importaci�n de Mercader�as </option>
				<option value=4>4-Exportaci�n Cereales, Aceites y Subproductos </option>				
				<option value=5>5-Exportaci�n Cereales, Aceites y Subproductos (Anual) </option>				
				<option value=7>7-Total Exportado por Destino - CAS </option>
				<option value=8>8-Detalle de Cargas por Sitio </option>				
				<option value=9>9-Porcentaje de Participaci�n por Sitio </option>				
				<option value=10>10-Participaci�n por Terminal </option>
				<option value=11>11-Exportaci�n de Pescado </option>
				<option value=12>12-Exportaci�n Inflamables</option>
				<option value=13>13-Cabotaje Mar�timo Nacional - Removido Salidas - Cargas</option>				
				<option value=14>14-Cabotaje Mar�timo Nacional - Removido Entradas - Descargas</option>				
				<option value=15>15-Removido Inflamables</option>				
				<option value=18>18-Removido, Importaci�n y Exportaci�n - P�troleo Crudo</option>				
				<option value=19>19-Movimientos de Buques por Sitio</option>								
				<option value=20>20-Movimiento General</option>				
				<option value=21>21-Detalle Atenci�n Buques por Agencia</option>
				<option value=22>22-Exportaci�n Frutas - Detalle de Buques</option>				
				</select>
		</td>					
		
	</tr>
	-->
	<tr valign="top" height="100%">
		<td colspan="6" align="center" width="100%">
      		<iframe name="ifrm" scrolling="Yes" src="" width="100%" height="100%"></iframe>
      	</td>
	</tr>
	<tr>
		<td colspan="6" height="10">
		</td>
	</tr>
</table>
</form>	
</body>
<script>
	Actualizar('ifrm');
</script>


</html>

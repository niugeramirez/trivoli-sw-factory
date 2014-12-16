<% Option Explicit %>
<!--#include virtual="/serviciolocal/shared/db/conn_db.inc"-->
<!--#include virtual="/serviciolocal/shared/inc/sec.inc"-->
<!--#include virtual="/serviciolocal/shared/inc/const.inc"-->
<!--#include virtual="/serviciolocal/shared/inc/fecha.inc"-->
<!--#include virtual="/serviciolocal/shared/inc/antigfec.inc"-->
<!--#include virtual="/serviciolocal/shared/inc/sqls.inc"-->
<!--
-----------------------------------------------------------------------------
Archivo: rep_log_seg_00.asp
Autor: Raul Chinestra
Creacion: 29/06/2006
Descripcion: Reporte de Log
 -----------------------------------------------------------------------------
-->
<html>
<head>
<link href="/serviciolocal/shared/css/tables3.css" rel="StyleSheet" type="text/css">
<title><%= Session("Titulo")%> Reporte de Log - Ticket </title>
<script src="/serviciolocal/shared/js/fn_windows.js"></script>
<script src="/serviciolocal/shared/js/fn_confirm.js"></script>
<script src="/serviciolocal/shared/js/fn_ayuda.js"></script>
<script src="/serviciolocal/shared/js/fn_fechas.js"></script>
<script src="/serviciolocal/shared/js/fn_ay_generica.js"></script>
<script>

<%
on error goto 0 
Dim l_rs
Dim l_sql
%>

function Imprimir(){
	parent.frames.ifrm.focus();
	window.print();
}


function Actualizar(destino){

	var param;	
	
	if (document.datos.fecini.value == "") {
  		alert("La Fecha Desde No debe ser Vac�a");
  		document.datos.fecini.focus();
		return;
	}

	if (document.datos.fecfin.value == "") {
  		alert("La Fecha Hasta No debe ser Vac�a");
  		document.datos.fecfin.focus();
		return;
	}

	if (!validarfecha(document.datos.fecini)) {
  		document.datos.fecini.focus();
		return;
	}
	
	if (!validarfecha(document.datos.fecfin)) {
    	document.datos.fecfin.focus();
		return;
	}

	if (!(menorque(document.datos.fecini.value,document.datos.fecfin.value))) {
		alert("La Fecha Desde debe ser menor o igual que la Fecha Hasta.");
		document.datos.fecini.focus();
		return;
	}		  	
	param = "qtiplognro=" + document.all.tiplognro.value + "&qlogusr=" + document.all.logusr.value;			
	param = param + "&qfecini=" + document.all.fecini.value + "&qfecfin=" + document.all.fecfin.value;			
	param = param + "&qcarpornum=" + document.all.carpornum.value;
	param = param + "&qmovcod=" + document.all.movcod.value;
	param = param + "&qlogdetest=" + document.all.logdetest.value;	
	param = param + "&qdet=" + document.all.det.checked;		
	if (destino== "exel")
    	abrirVentana("rep_log_seg_04.asp?" + param,'execl',250,150);
	else
		document.ifrm.location = "rep_log_seg_01.asp?"+ param;			
	
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

function HabDet(){
	
	if (document.all.det.checked) {
		document.datos.carpornum.className = "habinp"
		document.datos.movcod.className = "habinp"		
		document.datos.logdetest.className = "habinp"
		document.datos.carpornum.disabled = false		
		document.datos.movcod.disabled = false		
		document.datos.logdetest.disabled = false		
	}
	else {
		document.datos.carpornum.className = "deshabinp"
		document.datos.movcod.className = "deshabinp"		
		document.datos.logdetest.className = "deshabinp"			
		document.datos.carpornum.disabled = true
		document.datos.movcod.disabled = true	
		document.datos.logdetest.disabled = true
		document.datos.carpornum.value = ""
		document.datos.movcod.value = ""		
		document.datos.logdetest.value = 0		
	}
}


</script>

</head>
<body leftmargin="0" topmargin="0" rightmargin="0" bottommargin="0" onload="Javascript:document.datos.tiplognro.focus();" >
<form name="datos">
      <table border="0" cellpadding="0" cellspacing="0" height="100%">
        <tr style="border-color :CadetBlue;">
          <td align="left" class="barra" nowrap>
		  <!--<a class=sidebtnSHW href="Javascript:window.close();">Salir</a>--></td>
          <td align="right" class="barra" colspan="3">
 		  <a class=sidebtnSHW href="Javascript:Actualizar('ifrm')">Actualizar</a>		  
		  <a class=sidebtnSHW href="Javascript:Imprimir()">Imprimir</a>		  	  		  
		  <!-- <a class=sidebtnSHW href="Javascript:Actualizar('exel')">Excel</a> -->
		  &nbsp;
		  <a class=sidebtnHLP href="Javascript:ayuda('<%= Request.ServerVariables("SCRIPT_NAME")%>');">Ayuda</a>
		  </td>
        </tr>
		<tr>
			<td align="right">
				&nbsp;&nbsp;<b>Tipo Log:</b>
			</td>
			<td>
				<select name="tiplognro" size="1" style="width:200;">
					<option value=0 selected >&laquo; Todos los Tipos &raquo;</option>
					<%Set l_rs = Server.CreateObject("ADODB.RecordSet")
					l_sql = "SELECT tiplogdes, tiplognro "
					l_sql  = l_sql  & " FROM tkt_tipolog "
					l_sql  = l_sql  & " ORDER BY tiplogdes "
					rsOpen l_rs, cn, l_sql, 0
					do until l_rs.eof	%>	
					<option value=<%= l_rs("tiplognro") %> > 
					<%= l_rs("tiplogdes")%> </option>
					<%	l_rs.Movenext
					loop
					l_rs.Close %>
				</select>
			</td>
			<td align="right"><b>Usuario:</b>
			</td>
			<td>
				<select name="logusr" size="1" style="width:200;">
					<option value=0 selected >&laquo; Todos los Usuarios &raquo;</option>
					<%Set l_rs = Server.CreateObject("ADODB.RecordSet")
					l_sql = "SELECT iduser "
					l_sql  = l_sql  & " FROM user_per "
					l_sql  = l_sql  & " ORDER BY iduser "
					rsOpen l_rs, cn, l_sql, 0
					do until l_rs.eof	%>	
					<option value=<%= l_rs("iduser") %> > 
					<%= l_rs("iduser")%> </option>
					<%	l_rs.Movenext
					loop
					l_rs.Close %>
				</select>
			</td>
		</tr>
	<tr>
		<td align="right">
			<b>Fecha Desde:</b>
		</td>
		<td>
			<input  type="text" name="fecini" size="10" maxlength="10" value="" >
			<a href="Javascript:Ayuda_Fecha(document.datos.fecini);"><img src="/serviciolocal/shared/images/cal.gif" border="0"></a>
		</td>
		<td align="right">
			<b>Fecha Hasta:</b>
		</td>
		<td>			
			<input  type="text" name="fecfin" size="10" maxlength="10" value="">
				<a href="Javascript:Ayuda_Fecha(document.datos.fecfin);"><img src="/serviciolocal/shared/images/cal.gif" border="0"></a>
		</td>
	</tr>		
	<tr>
		<td align="right">
			<b>Detallado:</b>			
		</td>		
		<td align="left">
			<input   type="Checkbox" name="det" checked onclick="Javascript:HabDet();"  >
		</td>					
		<td align="right">
			<b>Estado:</b>
		</td>
		<td>
			<select name="logdetest" size="1" style="width:200;">
				<option value=0 selected >&laquo; Todos los Estados &raquo;</option>
				<option value="I" >I = Ingresado</option>
				<option value="R" >R = Rechazado</option>				
			</select>
		</td>
	</tr>		
	<tr>
		<td align="right">
			<b>Carta de Porte:</b>
		</td>
		<td >			
			<input  type="text" name="carpornum" size="15" maxlength="13" value="" > (0000-00000000)
		</td>
		<td align="right">
			<b>Movimiento:</b>
		</td>
		<td>					
			<input  type="text" name="movcod" size="15" maxlength="10" value=""> (10 caracteres)
		</td>
	</tr>				
	<tr valign="top" height="100%">
          <td colspan="4" align="center">
      	  <iframe name="ifrm" scrolling="Yes" src="" width="100%" height="100%"></iframe> 
	      </td>
    </tr>
    <tr>
          <td colspan="4" height="10">
	      </td>
    </tr>
	</table>
</form>	
</body>
<iframe name="valida" scrolling="Yes" src="" width="100%" height="100%"></iframe> 

</html>

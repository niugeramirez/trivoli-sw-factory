<% Option Explicit %>
<!--#include virtual="/serviciolocal/shared/db/conn_db.inc"-->
<!--#include virtual="/serviciolocal/shared/inc/sec.inc"-->
<!--#include virtual="/serviciolocal/shared/inc/const.inc"-->
<!--#include virtual="/serviciolocal/shared/inc/fecha.inc"-->
<!--#include virtual="/serviciolocal/shared/inc/antigfec.inc"-->
<!--#include virtual="/serviciolocal/shared/inc/sqls.inc"-->
<!--
-----------------------------------------------------------------------------
Archivo: rep_vessel_tender_con_00.asp
Autor: Raul Chinestra
Creacion: 01/02/2008
Descripcion: Reporte de Vessel Tender
 -----------------------------------------------------------------------------
-->
<html>
<head>
<link href="/serviciolocal/shared/css/tables3.css" rel="StyleSheet" type="text/css">
<title><%= Session("Titulo")%> Vessel Tender - buques </title>
<script src="/serviciolocal/shared/js/fn_windows.js"></script>
<script src="/serviciolocal/shared/js/fn_confirm.js"></script>
<script src="/serviciolocal/shared/js/fn_ayuda.js"></script>
<script src="/serviciolocal/shared/js/fn_fechas.js"></script>
<script src="/serviciolocal/shared/js/fn_ay_generica.js"></script>
<script>

<% on error goto 0 
Dim l_rs
Dim l_sql
 %>

function Imprimir(){
	parent.frames.ifrm.focus();
	window.print();	
}


function Actualizar(destino){

	var param;
	
	// Controlo que ingrese la fecha desde
	if ((document.datos.fecini.value == "")) {
  		alert("Debe ingresar la Fecha Desde");
  		document.datos.fecini.focus();
		return;
	}	
	// Controlo que ingrese la fecha hasta
	if ((document.datos.fecfin.value == "")) {	
  		alert("Debe ingresar la Fecha Hasta");
  		document.datos.fecfin.focus();
		return;
	}	
	
	// Controlo que ingrese las dos fechas
	if ((document.datos.fecini.value == "") && (document.datos.fecfin.value != "" )) {
  		alert("Debe ingresar la Fecha Desde o borrar la Fecha Hasta");
  		document.datos.fecini.focus();
		return;
	}
	// Controlo que ingrese las dos fechas	
	if ((document.datos.fecfin.value == "") && (document.datos.fecini.value != "" )) {	
  		alert("Debe ingresar la Fecha Hasta o borrar la Fecha Desde");
  		document.datos.fecfin.focus();
		return;
	}
	
	// Si las dos fechas fueron ingresadas, valido a las mismas	
	if ((document.datos.fecini.value != "") && (document.datos.fecfin.value != "" )) {
	
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
	}		
	
	// Controlo que haya seleccionado al menos un Grupo
	if (document.datos.quatypnro.value == "0") {
  		alert("Debe ingresar un Grupo");
  		document.datos.quatypnro.focus();
		return;
	}	

	param = "qfecini=" + document.all.fecini.value + "&qfecfin=" + document.all.fecfin.value + "&qquatypnro=" + document.all.quatypnro.value;			
	if (destino== "exel")
    	abrirVentana("rep_vessel_tender_con_04.asp?" + param,'execl',250,150);
	else
		document.ifrm.location = "rep_vessel_tender_con_01.asp?"+ param;			
	
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
<body leftmargin="0" topmargin="0" rightmargin="0" bottommargin="0" >
<form name="datos">
      <table border="0" cellpadding="0" cellspacing="0" height="100%">
        <tr style="border-color :CadetBlue;">
          <td align="left" class="barra" nowrap>
		  <!--<a class=sidebtnSHW href="Javascript:window.close();">Salir</a>--></td>
          <td align="right" class="barra" colspan="3">
 		  <a class=sidebtnSHW href="Javascript:Actualizar('ifrm')">Actualizar</a>		  
		  <a class=sidebtnSHW href="Javascript:Imprimir()">Imprimir</a>		  
		  <!-- <a class=sidebtnSHW href="Javascript:Actualizar('exel')">Excel</a>		  -->
		  &nbsp;
		  <a class=sidebtnHLP href="Javascript:ayuda('<%= Request.ServerVariables("SCRIPT_NAME")%>');">Ayuda</a>
		  </td>
        </tr>
		<tr valign="top" height="100%">
          <td colspan="4" align="center">
      	  <iframe name="ifrm" scrolling="Yes" src="datos_personal_con_02.asp" width="100%" height="100%"></iframe> 
	      </td>
        </tr>
        <tr>
          <td colspan="4" height="10">
	      </td>
        </tr>
	</table>
</form>	
</body>
</html>

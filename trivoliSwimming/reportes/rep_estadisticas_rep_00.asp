<% Option Explicit %>
<!--#include virtual="/trivoliSwimming/shared/inc/sec.inc"-->
<!--#include virtual="/trivoliSwimming/shared/inc/const.inc"-->
<!--#include virtual="/trivoliSwimming/shared/db/conn_db.inc"-->
<html>
<head>
<link href="/trivoliSwimming/ess/shared/css/tables_gray.css" rel="StyleSheet" type="text/css">
<title>Cantidades entre Fechas</title>
<script src="/trivoliSwimming/shared/js/fn_windows.js"></script>
<script src="/trivoliSwimming/shared/js/fn_confirm.js"></script>
<script src="/trivoliSwimming/shared/js/fn_ayuda.js"></script>
<script src="/trivoliSwimming/shared/js/fn_fechas.js"></script>
<script src="/trivoliSwimming/shared/js/fn_ay_generica.js"></script>

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

		
$( "#fechadesde" ).datepicker({
	showOn: "button",
	buttonImage: "/trivoliSwimming/shared/images/calendar1.png",
	buttonImageOnly: true
});

$( "#fechahasta" ).datepicker({
	showOn: "button",
	buttonImage: "/trivoliSwimming/shared/images/calendar1.png",
	buttonImageOnly: true
});

});
</script>
<!-- Final Datepicker -->

<script>

<% on error goto 0
Dim l_rs
Dim l_sql
Dim l_id
Dim l_idrecursoreservable

l_id = 0
l_idrecursoreservable = 0
%>

function Imprimir(){
	document.ifrm.focus();
	window.print();	
}

function Actualizar(destino){

	var param;
	//Fechas	
	
	if (document.datos.fechadesde.value == "")  {
  		alert("Debe ingresar la Fecha Desde ");
  		document.datos.fechadesde.focus();
		return;
	}
	
	if (document.datos.fechahasta.value == "")  {
  		alert("Debe ingresar la Fecha Hasta ");
  		document.datos.fechahasta.focus();
		return;
	}	
	
	if (!menorque(document.datos.fechadesde.value ,document.datos.fechahasta.value)){
		alert('La fecha hasta debe ser mayor que la fecha desde.');
  		document.datos.fechadesde.focus();
		return;
	}		

	param = "qfechadesde=" + document.all.fechadesde.value + "&qfechahasta=" + document.all.fechahasta.value + "&idrecursoreservable=" + document.all.idrecursoreservable.value; // + document.all.repnro.value;
	
	if (destino== "exel")
    	abrirVentana("rep_estadisticas_rep_01.asp?" + param + "&excel=true",'execl',250,150);
	else
		document.ifrm.location = "rep_estadisticas_rep_02.asp?" + param;			
	
}

function Ayuda_Fecha(txt){
	var jsFecha = Nuevo_Dialogo(window, '/trivoliSwimming/shared/js/calendar.html', 16, 15);
	if (jsFecha == null){
		//txt.value = '';
	}else{
		txt.value = jsFecha;
		//DiadeSemana(jsFecha);
	}
}



</script>

</head>
<body leftmargin="0" topmargin="0" rightmargin="0" bottommargin="0" onload="Javascript:document.datos.fecini.focus();" >
<form name="datos">
<table border="0" cellpadding="0" cellspacing="0" height="100%">
	<tr style="border-color :CadetBlue;">
		<td align="left" class="barra" nowrap>
			<!--<a class=sidebtnSHW href="Javascript:window.close();">Salir</a>--></td>
		<td align="right" class="barra" >
			<a class=sidebtnSHW href="Javascript:Actualizar('ifrm')"><img  src="/trivoliSwimming/shared/images/Buscar_24.png" border="0" title="Buscar"></a>		  
			<!--<a class=sidebtnSHW href="Javascript:Imprimir()">Imprimir</a>	-->	  
			<a class=sidebtnSHW href="Javascript:Actualizar('exel')"><img  src="/trivoliSwimming/shared/images/Excel-icon_24.png" border="0" title="Excel"></a> 
			&nbsp;
			
		</td>
	</tr>
		<tr>
			<td align="center" colspan="2">
				<table border="0">
					<input type="hidden" name="filtro" value="">

					<tr>
						<td align="right"><b>Fecha Desde: </b></td>
						<td><input id="fechadesde" type="text" name="fechadesde" size="10" maxlength="10" value="<%= "01/01/"&year(date())-1%>" >							
						</td>
						<td align="right"><b>Fecha Hasta: </b></td>
						<td><input  id="fechahasta" type="text" name="fechahasta" size="10" maxlength="10" value="<%= "31/12/"&year(date())%>" >							
						</td>						
						
						<td  align="right" nowrap><b>Indicador: </b></td>
						<td><select name="idrecursoreservable" size="1" style="width:300;">
								<option value=0 selected>Seleccione un Indicador</option>
								<option value=1 >Ventas Proyectadas Vs Reales</option>
								<option value=2 >Stock</option>
								<option value=3 >Instalaciones</option>
								<option value=4 >Caja</option> 
								<option value=5 >Caja/Reponsable</option>
								<option value=6 >Obligaciones a Pagar</option>
								<option value=7>Estado Cheques</option>
								<option value=8>Obligaciones a Cobrar</option>
								<option value=9>Utilidad Neta</option>
							</select>
							
						</td>			
																		
					</tr>	

				</table>
			</td>
		</tr>
        <tr valign="top" height="100%">
          <td colspan="2" style="" width="100%">
      	  <iframe scrolling="yes" name="ifrm" src="" width="100%" height="100%"></iframe> 
	      </td>
        </tr>		
</table>
</form>	
</body>
</html>

<% Option Explicit %>
<!--#include virtual="/turnos/shared/inc/sec.inc"-->
<!--#include virtual="/turnos/shared/inc/const.inc"-->
<!--#include virtual="/turnos/shared/db/conn_db.inc"-->
<html>
<head>
<link href="/turnos/ess/shared/css/tables_gray.css" rel="StyleSheet" type="text/css">
<title>Visitas entre Fechas</title>
<script src="/turnos/shared/js/fn_windows.js"></script>
<script src="/turnos/shared/js/fn_confirm.js"></script>
<script src="/turnos/shared/js/fn_ayuda.js"></script>
<script src="/turnos/shared/js/fn_fechas.js"></script>
<script src="/turnos/shared/js/fn_ay_generica.js"></script>

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
	buttonImage: "/turnos/shared/images/calendar1.png",
	buttonImageOnly: true
});

$( "#fechahasta" ).datepicker({
	showOn: "button",
	buttonImage: "/turnos/shared/images/calendar1.png",
	buttonImageOnly: true
});

});
</script>
<!-- Final Datepicker -->
<!-- Inicio MULTIPLE SELECCION -->	
<script src="../js/jquery.sumoselect.js"></script>
<link href="../js/sumoselect.css" rel="stylesheet" />

<script type="text/javascript">
    $(document).ready(function () {
        window.asd = $('.SlectBox').SumoSelect({ csvDispCount: 3 });
        window.test = $('.testsel').SumoSelect({okCancelInMulti:true });
        window.testSelAll = $('.testSelAll').SumoSelect({okCancelInMulti:true, selectAll:true });
        window.testSelAll2 = $('.testSelAll2').SumoSelect({selectAll:true });
    });
</script>
<style type="text/css">
    body{font-family: 'Helvetica Neue', Helvetica, Arial, sans-serif;color:#444;font-size:13px;}
    p,div,ul,li{padding:0px; margin:0px;}
</style>
<!-- Final MULTIPLE SELECCION -->
<script>

<% on error goto 0
Dim l_rs
Dim l_sql
Dim l_id


l_id = 0
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
	
	// Obras Sociales
    lista = document.all.idos;
	opciones = lista.options
    osociales = '';
    
	for (i=0;i<opciones.length;i++) {
         if (opciones[i].selected == true ) {
		     if  (osociales == '' ) {
			 	osociales = opciones[i].value;
			 }	
			 else {
			 	osociales = osociales + ',' + opciones[i].value ;
			 }           
		
         }
    }	
	
	if (osociales == '' ){
		alert('Debe Seleccionar al menos una Obra Social.');
		return;
		
	}	
	
	if ( document.datos.afiliado_oblig.checked == false && document.datos.afiliado_volunt.checked == false ){
	alert("Debe seleccionar si el Reporte se genera para Afiliados Obligatorios y/o Voluntarios.");
	document.datos.afliliado_oblig.focus();
	return;
	}

	osociales = "(" + osociales + ")";
	
	tipo = 'T'
	
	if ( document.datos.afiliado_oblig.checked == false ){
		tipo = 'V'
	}
	
	if ( document.datos.afiliado_volunt.checked == false ){
		tipo = 'O'
	}
	
	param = "qfechadesde=" + document.all.fechadesde.value + "&qfechahasta=" + document.all.fechahasta.value + "&idos=" + osociales + "&tipo=" + tipo; 
	
	if (destino== "exel"){
    	abrirVentana("rep_visitas_por_OS_rep_01.asp?" + param + "&excel=true",'excel',250,150);
	}
	else
	{
		if (destino== "txt"){
			abrirVentana("rep_visitas_por_OS_rep_02.asp?" + param,'txt',600,120);
		} 
		else
		{
			document.ifrm.location = "rep_visitas_por_OS_rep_01.asp?" + param;			
		}
	}
}

function Ayuda_Fecha(txt){
	var jsFecha = Nuevo_Dialogo(window, '/turnos/shared/js/calendar.html', 16, 15);
	if (jsFecha == null){
		//txt.value = '';
	}else{
		txt.value = jsFecha;
		//DiadeSemana(jsFecha);
	}
}



</script>

</head>
<body leftmargin="0" topmargin="0" rightmargin="0" bottommargin="0" onload="Javascript:document.datos.fechadesde.focus();" >
<form name="datos">
<table border="0" cellpadding="0" cellspacing="0" height="100%">
	<tr style="border-color :CadetBlue;">
		<td align="left" class="barra" nowrap>
			<!--<a class=sidebtnSHW href="Javascript:window.close();">Salir</a>--></td>
		<td align="right" class="barra" >
			<a class=sidebtnSHW href="Javascript:Actualizar('ifrm')"><img  src="/turnos/shared/images/Buscar_24.png" border="0" title="Buscar"></a>		  
			<!--<a class=sidebtnSHW href="Javascript:Imprimir()">Imprimir</a>	-->	  
			<a class=sidebtnSHW href="Javascript:Actualizar('exel')"><img  src="/turnos/shared/images/Excel-icon_24.png" border="0" title="Excel"></a> 
			<a class=sidebtnSHW href="Javascript:Actualizar('txt')"><img  src="/turnos/shared/images/txt.png" border="0" title="Generar TXT"></a> 
			&nbsp;
		</td>
	</tr>
		<tr>
			<td align="center" colspan="2">
				<table border="0">
					<input type="hidden" name="filtro" value="">
		

					<tr>
						<td  align="right" nowrap><b>Fecha Desde:</b></td>
						<td><input id="fechadesde" type="text" name="fechadesde" value=""></td>
						
						<td  align="right" nowrap><b>Fecha Hasta:</b></td>
						<td align="right"><input id="fechahasta" type="text" name="fechahasta"></td>
												
						
						<td  align="right" nowrap><b>Obra Social: </b></td>
						<td colspan="3"><select name="idos" multiple="multiple" placeholder="Obras Sociales" onchange="console.log($(this).children(':selected').length)" class="testSelAll">
							<%Set l_rs = Server.CreateObject("ADODB.RecordSet")
								l_sql = "SELECT  * "
								l_sql  = l_sql  & " FROM obrassociales  "
							    l_sql =  l_sql & " where obrassociales.empnro = " & Session("empnro") 
								l_sql  = l_sql  & " ORDER BY descripcion "
								rsOpen l_rs, cn, l_sql, 0
								do until l_rs.eof		%>	
								<option value= <%= l_rs("id") %> > 
								<%= l_rs("descripcion") %> </option>
								<%	l_rs.Movenext
								loop
								l_rs.Close %>
							</select>
						</td>							
						
						<td align="right"><b>Obligatorios:</b></td>
						<td>
							<input  type=checkbox name="afiliado_oblig">
						</td>
						
						<td align="right"><b>Voluntarios:</b></td>
						<td>
							<input  type=checkbox name="afiliado_volunt">
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

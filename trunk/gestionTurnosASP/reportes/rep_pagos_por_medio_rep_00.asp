<% Option Explicit %>
<!--#include virtual="/turnos/shared/inc/sec.inc"-->
<!--#include virtual="/turnos/shared/inc/const.inc"-->
<!--#include virtual="/turnos/shared/db/conn_db.inc"-->
<html>
<head>
<link href="/turnos/ess/shared/css/tables_gray.css" rel="StyleSheet" type="text/css">
<title>Pagos entre Fechas</title>
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

	
	param = "qfechadesde=" + document.all.fechadesde.value + "&qfechahasta=" + document.all.fechahasta.value + "&idmedio=" + document.all.id.value + "&idrecursoreservable=" + document.all.idrecursoreservable.value; // + document.all.repnro.value;
	
	if (destino== "exel")
    	abrirVentana("rep_pagos_por_medio_rep_01.asp?" + param + "&excel=true",'execl',250,150);
	else
		document.ifrm.location = "rep_pagos_por_medio_rep_01.asp?" + param;			
	
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
<body leftmargin="0" topmargin="0" rightmargin="0" bottommargin="0" onload="Javascript:document.datos.fecini.focus();" >
<form name="datos">
<table border="0" cellpadding="0" cellspacing="0" height="100%">
	<tr style="border-color :CadetBlue;">
		<td align="left" class="barra" nowrap>
			<!--<a class=sidebtnSHW href="Javascript:window.close();">Salir</a>--></td>
		<td align="right" class="barra" >
			<a class=sidebtnSHW href="Javascript:Actualizar('ifrm')">Actualizar</a>		  
			<!--<a class=sidebtnSHW href="Javascript:Imprimir()">Imprimir</a>	-->	  
			<a class=sidebtnSHW href="Javascript:Actualizar('exel')">Excel</a> 
			&nbsp;
			
		</td>
	</tr>
		<tr>
			<td align="center" colspan="2">
				<table border="0">
					<input type="hidden" name="filtro" value="">

					<tr>
						<td align="right"><b>Fecha Desde: </b></td>
						<td><input id="fechadesde" type="text" name="fechadesde" size="10" maxlength="10" value="<%= date()%>" >							
						</td>
						<td align="right"><b>Fecha Hasta: </b></td>
						<td><input  id="fechahasta" type="text" name="fechahasta" size="10" maxlength="10" value="<%= date()%>" >							
						</td>						
						<td  align="right" nowrap><b>Medio de Pago: </b></td>
						<td><select name="id" size="1" style="width:200;">
								<option value=0 selected>Todos los Medios</option>
								<%Set l_rs = Server.CreateObject("ADODB.RecordSet")
								l_sql = "SELECT  * "
								l_sql  = l_sql  & " FROM mediosdepago  "
								l_sql  = l_sql  & " ORDER BY titulo "
								rsOpen l_rs, cn, l_sql, 0
								do until l_rs.eof		%>	
								<option value= <%= l_rs("id") %> > 
								<%= l_rs("titulo") %> </option>
								<%	l_rs.Movenext
								loop
								l_rs.Close %>
							</select>
							<script>document.datos.id.value= "<%= l_id %>"</script>
						</td>	
						<td  align="right" nowrap><b>M&eacute;dico: </b></td>
						<td><select name="idrecursoreservable" size="1" style="width:200;">
								<option value=0 selected>Todos los M&eacute;dico</option>
								<%Set l_rs = Server.CreateObject("ADODB.RecordSet")
								l_sql = "SELECT  * "
								l_sql  = l_sql  & " FROM recursosreservables  "
								l_sql  = l_sql  & " ORDER BY descripcion "
								rsOpen l_rs, cn, l_sql, 0
								do until l_rs.eof		%>	
								<option value= <%= l_rs("id") %> > 
								<%= l_rs("descripcion") %> </option>
								<%	l_rs.Movenext
								loop
								l_rs.Close %>
							</select>
							<script>document.datos.idrecursoreservable.value= "<%= l_idrecursoreservable %>"</script>
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

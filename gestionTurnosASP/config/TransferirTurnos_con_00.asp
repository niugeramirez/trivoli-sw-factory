<% Option Explicit %>
<!--#include virtual="/turnos/shared/inc/sec.inc"-->
<!--#include virtual="/turnos/shared/inc/const.inc"-->
<!--#include virtual="/turnos/shared/db/conn_db.inc"-->
<% 
'Archivo: transferirturnos_con_00.asp
'Descripción: Transferir Turnos
'Autor : Raul Chinestra
'Fecha: 01/07/2015


on error goto 0

  Dim l_rs
  Dim l_sql
  
  Dim l_hd
  Dim l_md
  Dim l_hh
  Dim l_mh  
  Dim l_cabnro 
  
  l_cabnro = request("cabnro")
  
%>
<html>
<head>
<link href="/turnos/ess/shared/css/tables_gray.css" rel="StyleSheet" type="text/css">
<!--<link href="/turnos/shared/css/tables_gray.css" rel="StyleSheet" type="text/css">-->
<title>Asignar Turnos</title>
<script src="/turnos/shared/js/fn_windows.js"></script>
<script src="/turnos/shared/js/fn_confirm.js"></script>
<script src="/turnos/shared/js/fn_ayuda.js"></script>
<script src="/turnos/shared/js/fn_fechas.js"></script>



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


function Buscar(){
	var tieneotro;
	var estado;
	document.datos.filtro.value = "";
	tieneotro = "no";
	estado = "si";
	
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
	
	

	// fec. desde
	if (document.datos.fechadesde.value != ""){
		if (tieneotro == "si"){
			document.datos.filtro.value += " AND " ;
			tieneotro = "si";
		}
		if (validarfecha(document.datos.fechadesde)){
			document.datos.filtro.value += " CONVERT(VARCHAR(10), calendarios.fechahorainicio, 101)  >= " + cambiafecha(document.datos.fechadesde.value,true,1) + "";
			tieneotro = "si";
		}else{
			estado = "no";
		}
	}
	
	// fec. hasta
	if (document.datos.fechahasta.value != ""){
		if (tieneotro == "si"){
			document.datos.filtro.value += " AND " ;
			tieneotro = "si";
		}
		if (validarfecha(document.datos.fechahasta)){
			document.datos.filtro.value += " CONVERT(VARCHAR(10), calendarios.fechahorainicio, 101)  <= " + cambiafecha(document.datos.fechahasta.value,true,1) + "";
			tieneotro = "si";
		}else{
			estado = "no";
		}
	}	
	
	// Dia
    lista = document.all.dia;
	opciones = lista.options
    DiasSemana = '';
    for (i=0;i<opciones.length;i++) {
         if (opciones[i].selected == true ) {
		     if  (DiasSemana == '' ) {
			 	DiasSemana = opciones[i].value;
			 }	
			 else {
			 	DiasSemana = DiasSemana + ',' + opciones[i].value ;
			 }           
		
         }
    }		
	
	if (DiasSemana == '' ){
		alert('Debe Seleccionar al menos un Dia.');
		return;
		
	}		
	
	// Medicos
    lista = document.all.Med;
	opciones = lista.options
    medicos = '';
    for (i=0;i<opciones.length;i++) {
         if (opciones[i].selected == true ) {
		     if  (medicos == '' ) {
			 	medicos = opciones[i].value;
			 }	
			 else {
			 	medicos = medicos + ',' + opciones[i].value ;
			 }           
		
         }
    }	
	
	if (medicos == '' ){
		alert('Debe Seleccionar al menos un Medico.');
		return;
		
	}		
	
	
	if (document.datos.id.value != 0){
		if (tieneotro == "si"){
			document.datos.filtro.value += " AND calendarios.idrecursoreservable in (" + medicos + ")";
		}else{
			document.datos.filtro.value += " calendarios.idrecursoreservable in (" + medicos + ")";
		}
		tieneotro = "si";
	}	
	
	if (estado == "si"){
		window.ifrm.location = 'transferirturnos_con_01.asp?cabnro=' + document.datos.cabnro.value + "&fechadesde="+document.datos.fechadesde.value + "&fechahasta="+document.datos.fechahasta.value + "&hd="+document.datos.hd.value + "&md="+ document.datos.md.value + "&hh="+ document.datos.hh.value + "&mh="+ document.datos.mh.value + "&diassemana="+ DiasSemana + '&filtro=' + document.datos.filtro.value;
	}
}


function Limpiar(){

	document.datos.fechadesde.value = "<%= date() - 1 %>";
	document.datos.fechahasta.value = "<%= date() %>";

	document.datos.sernro.value     = 0;
	document.datos.pronro.value     = 0;	
/*	
	document.datos.pronro.value     = 0;
	document.datos.ternro.value     = 0;
	document.datos.pornro.value     = 0;
	document.datos.clinro.value     = 0;	
	document.datos.ctrnum.value     = 0;	
	document.datos.txtctrnum.value  = "";	
*/	
	window.ifrm.location = 'buques_con_01.asp';
}



</script>
</head>
<body leftmargin="0" topmargin="0" rightmargin="0" bottommargin="0" onload="Javascript:document.datos.fechadesde.focus();">
      <table border="0" cellpadding="0" cellspacing="0" height="100%" width="100%">
        <tr style="border-color :CadetBlue;">
          <td align="left" class="barra">&nbsp;</td>
          <td nowrap align="right" class="barra">        
		  &nbsp;  
		  </td>
        </tr>		
		<tr>
			<td align="center" colspan="2">
				<table border="0">
					<form name="datos">
					<input type="hidden" name="filtro" value="">

					<tr>
						<td align="right"><b>Fecha Desde: </b></td>
						<td><input  id="fechadesde" type="text" name="fechadesde" size="10" maxlength="10" value="<%= date()%>" >		
						
						<td align="right"><b>Fecha Hasta: </b></td>
						<td><input  id="fechahasta" type="text" name="fechahasta" size="10" maxlength="10" value="<%= date() + 7%>" >			
						
					</tr>						
					<tr>
						<td  align="right" nowrap><b>Desde: </b></td>
						<td ><select name="hd" size="1" style="width:50;">
								<%
								l_hd = 0  
								do while clng(l_hd) < 24 %>
								<option value= <%= right("0" & l_hd, 2) %>> <%= right("0" & l_hd, 2) %> </option>
								<%	l_hd = clng(l_hd) + 1
								loop
								%>
							</select>							
						    <b>:</b>
							<select name="md" size="1" style="width:50;">
								<%
								l_md = 0  
								do while clng(l_md) < 60 %>
								<option value= <%= right("0" & l_md, 2) %>> <%= right("0" & l_md, 2) %> </option>
								<%	l_md = clng(l_md) + 15
								loop
								%>
							</select>							
						</td>			
						<td  align="right" nowrap><b>Hasta: </b></td>
						<td ><select name="hh" size="1" style="width:50;">
								<%
								l_hh = 0  
								do while clng(l_hh) < 24 %>
								<option value= <%= right("0" & l_hh, 2) %>> <%= right("0" & l_hh, 2) %> </option>
								<%	l_hh = clng(l_hh) + 1
								loop
								%>
							</select>	
							<script>document.datos.hh.value="23"</script>							
						<b>:</b>
						   <select name="mh" size="1" style="width:50;">
								<%
								l_mh = 0  
								do while clng(l_mh) < 60 %>
								<option value= <%= right("0" & l_mh, 2) %>> <%= right("0" & l_mh, 2) %> </option>
								<%	l_mh = clng(l_mh) + 15
								loop
								%>
							</select>							
						</td>	
					</tr>	
					<tr>
						<td  align="right" nowrap><b>Dia Semana: </b></td>
						<td colspan="3"><select name="dia" multiple="multiple" placeholder="dia" onchange="console.log($(this).children(':selected').length)" class="testSelAll">
							<option value= "2" selected>Lunes </option>
							<option value= "3" selected>Martes </option>
							<option value= "4" selected>Miercoles </option>
							<option value= "5" selected>Jueves </option>
							<option value= "6" selected>Viernes </option>
							
						</td>					
					</tr>					
					<tr>
						<td  align="right" nowrap><b>M&eacute;dico: </b></td>
						<td colspan="3"><select name="Med" multiple="multiple" placeholder="Medicos" onchange="console.log($(this).children(':selected').length)" class="testSelAll">
							<%Set l_rs = Server.CreateObject("ADODB.RecordSet")
							l_sql = "SELECT  * "
							l_sql  = l_sql  & " FROM recursosreservables  "
							l_sql  = l_sql  & " WHERE recursosreservables.empnro = " & Session("empnro")
							l_sql  = l_sql  & " ORDER BY descripcion "
							rsOpen l_rs, cn, l_sql, 0
							do until l_rs.eof		%>	
							<option value= "<%= l_rs("id") %>" > 
							<%= l_rs("descripcion") %> </option>
							<%	l_rs.Movenext
							loop
							l_rs.Close %>
						</td>					
					</tr>						
					<tr>
						<td align="center" colspan="4">	
								
							<!--<td ><img src="../shared/images/gen_rep/boton_01.gif" width="5.9"></td>-->
							<a class="sidebtnABM" href="Javascript:Buscar();"><img  src="/turnos/shared/images/Buscar_24.png" border="0" title="Buscar"></a>
							<!--<td  background="../shared/images/gen_rep/boton_05.gif"><img src="../shared/images/gen_rep/boton_03.gif" height="15"></td>-->
							<!--<a class="sidebtnABM" href="Javascript:Limpiar();"><img  src="/turnos/shared/images/Limpiar_24.png" border="0" title="Limpiar"></a>
							<!--<td ><img src="../shared/images/gen_rep/boton_06.gif"></td>-->
					</tr>
											

				</table>
			</td>
		</tr>		
		
        <tr valign="top" height="100%">
          <td colspan="2" style="" width="100%">
      	  <iframe scrolling="yes" name="ifrm" src="" width="100%" height="100%"></iframe> 
	      </td>
        </tr>		
		<input type="hidden" name="cabnro" value="<%= l_cabnro %>">
			</form>		
      </table>
</body>

<script>
	//Buscar();
</script>
</html>

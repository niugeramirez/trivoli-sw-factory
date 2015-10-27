<% Option Explicit %>
<!--#include virtual="/turnos/shared/inc/sec.inc"-->
<!--#include virtual="/turnos/shared/inc/const.inc"-->
<!--#include virtual="/turnos/shared/db/conn_db.inc"-->
<% 
'Archivo: contracts_con_00.asp
'Descripción: ABM de Contracts
'Autor : Raul Chinestra
'Fecha: 27/11/2007

' Son las listas de parametros a pasarle a los programas de filtro y orden
' En las mismas se deberan poner los valores, separados por un punto y coma

on error goto 0

' Filtro
  Dim l_Etiquetas  ' Son los nombres que deben aparecer en la ventana para que el usuario seleccione
  Dim l_Campos     ' Son los campos de la base que apareceran en la clausula where, que deben estar asociados a las etiquetas
  Dim l_Tipos      ' Son los tipos de datos que tienen los campos (N=Numerico, T=Texto y F=Fecha)

' Orden
  Dim l_Orden      ' Son las etiquetas que aparecen en el orden
  Dim l_CamposOr   ' Son los campos para el orden
  
' Filtro
  l_etiquetas = "Descripción:;Area"
  l_Campos    = "coudes;aredes"
  l_Tipos     = "T;T"

' Orden
  l_Orden     = "P/S:;Ctr Num:;Date:;Client:;Quality:;Volumen:;Port:;Term:;Company:;Product:"
  l_CamposOr  = "conpursal;ctrnum;confec;clidesabr;quadesabr;conquantity;pordes;terdes;comdesabr;prodesabr"

  Dim l_rs
  Dim l_sql
  
  Dim l_dia
  Dim l_mes
  Dim l_anio
  Dim l_id
  Dim l_fecha
  
l_dia = Request.Querystring("day")  
l_mes = Request.Querystring("Month")
l_anio = Request.Querystring("Year")
l_id = Request.Querystring("id")

If IsEmpty(l_dia) then 
	l_fecha = date()
else
	l_fecha = cstr(l_dia) & "/" & cstr(l_mes) & "/" & cstr(l_anio)
end if 
  
If IsEmpty(l_id) then   
	l_id = 0
end if
%>
<html>
<head>
<link href="/turnos/ess/shared/css/tables_gray.css" rel="StyleSheet" type="text/css">
<!--<link href="/turnos/shared/css/tables_gray.css" rel="StyleSheet" type="text/css">-->
<title>Generar Calendarios</title>
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

<script>

function orden(pag){
	abrirVentana('../shared/asp/orden_browse.asp?pagina='+pag+'&lista=<%= l_orden %>&campos=<%= l_camposOr%>&filtro='+escape(document.ifrm.datos.filtro.value),'',350,160)
}

function filtro(pag){
	abrirVentana('../shared/asp/filtro_browse.asp?pagina='+pag+'&campos=<%= l_campos%>&tipos=<%=l_tipos%>&etiquetas=<%=l_etiquetas%>&orden='+document.ifrm.datos.orden.value,'',250,160);
}

function llamadaexcel(){ 
	if (filtro == "")
		Filtro(true);
	else
		abrirVentana("contracts_con_excel.asp?orden=" + document.ifrm.datos.orden.value + "&filtro=" + escape(document.ifrm.datos.filtro.value),'execl',250,150);
}


function Buscar(generar){
	var tieneotro;
	var estado;
	document.datos.filtro.value = "";
	tieneotro = "no";
	estado = "si";

	// fec. desde
	if (document.datos.fechadesde.value != ""){
		if (tieneotro == "si"){
			document.datos.filtro.value += " AND " ;
			tieneotro = "si";
		}
		if (validarfecha(document.datos.fechadesde)){
			//Eugenio 18/09/2015, el formato de datetime con el filtro como estaba hacía que las fechas se comparen como strings y por ejemplo traiga registros de 2016 cuando no existen en la base 
			//document.datos.filtro.value += " CONVERT(VARCHAR(10), calendarios.fechahorainicio, 101)  >= " + cambiafecha(document.datos.fechadesde.value,true,1) + "";
			document.datos.filtro.value += " calendarios.fechahorainicio  >= Convert(datetime,  " + cambiafecha(document.datos.fechadesde.value,true,1) + ")";
			//FIN Eugenio 18/09/2015 
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
			//Eugenio 18/09/2015, el formato de datetime con el filtro como estaba hacia que las fechas se comparen como strings y por ejemplo traiga registros de 2016 cuando no existen en la base 
			//document.datos.filtro.value += " CONVERT(VARCHAR(10), calendarios.fechahorainicio, 101)  <= " + cambiafecha(document.datos.fechahasta.value,true,1) + "";			
			document.datos.filtro.value += " calendarios.fechahorainicio <=DATEADD(day, 1, Convert(datetime, " + cambiafecha(document.datos.fechahasta.value,true,1) + "))";													
			//FIN Eugenio 18/09/2015 
		}else{
			estado = "no";
		}
	}
	if (!menorque(document.datos.fechadesde.value ,document.datos.fechahasta.value)){
		alert('La fecha hasta debe ser mayor que la fecha desde.');
		estado = "no";
	}	
	/*
	// Servicio Local
	if (document.datos.sernro.value != 0){
		if (tieneotro == "si"){
			document.datos.filtro.value += " AND ser_legajo.legpar1 = '" + document.datos.sernro.value + "'";
		}else{
			document.datos.filtro.value += " ser_legajo.legpar1 = '" + document.datos.sernro.value + "'";
		}
		tieneotro = "si";
	}	
	*/
	
	if (document.datos.id.value == "0"){
		alert("Debe ingresar un Medico.");
		document.datos.id.focus();
		return;
	}	
	
	
	if (document.datos.id.value != 0){
		if (tieneotro == "si"){
			document.datos.filtro.value += " AND calendarios.idrecursoreservable = " + document.datos.id.value + "";
		}else{
			document.datos.filtro.value += " calendarios.idrecursoreservable = " + document.datos.id.value + "";
		}
		tieneotro = "si";
	}	
	/*	
	// Apellido
	if (document.datos.legape.value != 0){
		if (tieneotro == "si"){
			document.datos.filtro.value += " AND ser_legajo.legape like '*" + document.datos.legape.value + "*'";
		}else{
			document.datos.filtro.value += " ser_legajo.legape like '*" + document.datos.legape.value + "*'";
		}
		tieneotro = "si";
	}				
	// DNI
	if (document.datos.legdni.value != 0){
		if (tieneotro == "si"){
			document.datos.filtro.value += " AND ser_legajo.legdni = '" + document.datos.legdni.value + "'";
		}else{
			document.datos.filtro.value += " ser_legajo.legdni = '" + document.datos.legdni.value + "'";
		}
		tieneotro = "si";
	}					
	// Domicilio
	if (document.datos.legdom.value != 0){
		if (tieneotro == "si"){
			document.datos.filtro.value += " AND ser_legajo.legdom like '*" + document.datos.legdom.value + "*'";
		}else{
			document.datos.filtro.value += " ser_legajo.legdom like '*" + document.datos.legdom.value + "*'";
		}
		tieneotro = "si";
	}		
	// Medida de Proteccion
	if (document.datos.mednro.value != 0){
		if (tieneotro == "si"){
			document.datos.filtro.value += " AND ser_legajo.mednro = '" + document.datos.mednro.value + "'";
		}else{
			document.datos.filtro.value += " ser_legajo.mednro = '" + document.datos.mednro.value + "'";
		}
		tieneotro = "si";
	}							

	//alert(document.datos.filtro.value);
	*/
	if (estado == "si"){
		window.ifrm.location = 'GenerarCalendarios_con_01.asp?generar=' + generar + '&fechadesde='+ document.datos.fechadesde.value + '&fechahasta=' + document.datos.fechahasta.value + '&id=' + document.datos.id.value + '&filtro=' + document.datos.filtro.value;
	}
}


function AltaManual(){
	var tieneotro;
	var estado;
	document.datos.filtro.value = "";
	tieneotro = "no";
	estado = "si";

	
	if (document.datos.id.value == "0"){
		alert("Debe ingresar un Medico.");
		document.datos.id.focus();
		return;
	}	
	
     abrirVentana('calendarios_con_02.asp?id='+document.datos.id.value +'&Tipo=A','',550,350);
	
	
}



function Nuevo_Dialogo(w_in, pagina, ancho, alto)
{
 return w_in.showModalDialog(pagina,'', 'center:yes;dialogWidth:' + ancho.toString() + ';dialogHeight:' + alto.toString() + ';');
}
function Ayuda_Fecha(txt)
{
 var jsFecha = Nuevo_Dialogo(window, '/turnos/shared/js/calendar.html', 16, 15);

 if (jsFecha == null) txt.value = ''
 else txt.value = jsFecha;
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

function Contenido(){ 
	if (document.ifrm.datos.cabnro.value == 0) {
		alert("Debe seleccionar un Buque")
		return;
	}		
	else {
		abrirVentana("contenidos_con_00.asp?buqnro=" + document.ifrm.datos.cabnro.value,'',780,580);	
	}		
	
}

function fnctrnum(valor){
	if (valor == 0){
		document.datos.txtctrnum.disabled = true;
		document.datos.txtctrnum.className = "deshabinp";
	}else{
		document.datos.txtctrnum.disabled = false;
		document.datos.txtctrnum.className = "habinp";
	}
}

function TotalVolumen(valor){
	document.datos.totvol.value =  valor;
}


</script>
</head>
<body leftmargin="0" topmargin="0" rightmargin="0" bottommargin="0" onload="Javascript:document.datos.fechadesde.focus();">
      <table border="0" cellpadding="0" cellspacing="0" height="100%" width="100%">
        <tr style="border-color :CadetBlue;">
          <td align="left" class="barra">&nbsp;</td>
          <td nowrap align="right" class="barra">
		  
		  <%'eugenio 29/06/2015, unificacion de iconos  call MostrarBoton ("sidebtnABM", "Javascript:AltaManual();","Alta Manual de Calendario")%>	
		  <a class="sidebtnABM" href="Javascript:AltaManual();" ><img  src="/turnos/shared/images/Agregar_24.png" border="0" title="Alta">
				  
          <%' call MostrarBoton ("opcionbtn", "Javascript:abrirVentana('legajos_con_02.asp?Tipo=A','',800,500);","Alta")%>
		  &nbsp;&nbsp;
          <%' call MostrarBoton ("opcionbtn", "Javascript:eliminarRegistro(document.ifrm,'legajos_con_04.asp?cabnro=' + document.ifrm.datos.cabnro.value);","Baja")%>
		  &nbsp;&nbsp;
          <%' call MostrarBoton ("opcionbtn", "Javascript:abrirVentanaVerif('legajos_con_02.asp?Tipo=M&cabnro=' + document.ifrm.datos.cabnro.value,'',800,500);","Modifica")%>
		  &nbsp;&nbsp;
          <%' call MostrarBoton ("sidebtnSHW", "Javascript:llamadaexcel();","Excel")%>
		  &nbsp;&nbsp;
          <%' call MostrarBoton ("opcionbtn", "Javascript:Contenido();","Contenido")%>		  
		  		  
          <%' call MostrarBoton ("opcionbtn", "Javascript:orden('../../config/contracts_con_01.asp','',490,300);","Orden")%>		  
		  <!--
		  <a class=sidebtn href="Javascript:orden('../../config/contracts_con_01.asp');">Orden</a>
		  -->
		  &nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;
		  &nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;
		  &nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;
		  &nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;		  
		  </td>
        </tr>		
		<tr>
			<td align="center" colspan="2">
				<table border="0">
					<form name="datos">
					<input type="hidden" name="filtro" value="">
					<tr>
						<td align="right"><b>Fecha Desde: </b></td>
						<td><input  type="text" id="fechadesde" name="fechadesde" size="10" maxlength="10" value="<%= l_fecha%>" >
							
						</td>
						
						<td align="right"><b>Fecha Hasta: </b></td>
						<td><input  type="text" id="fechahasta" name="fechahasta" size="10" maxlength="10" value="<%= l_fecha %>" >							
						</td>					
						
						<td  align="right" nowrap><b>M&eacute;dico: </b></td>
						<td><select name="id" size="1" style="width:200;">
								<option value=0 selected>Seleccionar un M&eacute;dico</option>
								<%Set l_rs = Server.CreateObject("ADODB.RecordSet")
								l_sql = "SELECT  * "
								l_sql  = l_sql  & " FROM recursosreservables  "
								l_sql = l_sql & " where empnro = " & Session("empnro")
								l_sql  = l_sql  & " ORDER BY descripcion "
								rsOpen l_rs, cn, l_sql, 0
								do until l_rs.eof		%>	
								<option value= <%= l_rs("id") %> > 
								<%= l_rs("descripcion") %> </option>
								<%	l_rs.Movenext
								loop
								l_rs.Close %>
							</select>
							<script>document.datos.id.value= "<%= l_id %>"</script>
						</td>			
						

										<td ><a class="sidebtnABM" href="Javascript:Buscar(1);" ><img  src="/turnos/shared/images/calendar-add-icon_24.png" border="0" title="Generar Calendarios"></a></td>
										
										<td ><a class="sidebtnABM" href="Javascript:Buscar(0);"><img  src="/turnos/shared/images/Buscar_24.png" border="0" title="Buscar"></a></td>
										<!--<td ><img src="../shared/images/gen_rep/boton_06.gif"></td>-->						
								
					</tr>



					<tr>
					<!--					
						<td align="right"><b>Ctr. Nmbr: </b></td>
						<td><select name="ctrnum" size="1" onchange="fnctrnum(document.datos.ctrnum.value)">
								<option value=0 >Todos</option>
								<option value=1 >Comienza con</option>
								<option value=2 >Contiene</option>
								<option value=3 >Igual a</option>
							<input type="Text" name="txtctrnum" maxlength="20" style="width:100" disabled class="deshabinp">
							</select>
							<script> document.datos.ctrnum.value= "0"</script>
						</td>
						-->

						
					</tr>					

				</table>
			</td>
		</tr>		
		
        <tr valign="top" height="100%">
          <td colspan="2" style="" width="100%">
      	  <iframe scrolling="yes" name="ifrm" src="" width="100%" height="100%"></iframe> 
	      </td>
        </tr>		
			</form>		
      </table>
</body>

<script>
	//Buscar();
</script>
</html>

<% Option Explicit %>
<!--#include virtual="/serviciolocal/shared/inc/sec.inc"-->
<!--#include virtual="/serviciolocal/shared/inc/const.inc"-->
<!--#include virtual="/serviciolocal/shared/db/conn_db.inc"-->
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
  
%>
<html>
<head>
<link href="/serviciolocal/shared/css/tables4.css" rel="StyleSheet" type="text/css">
<!--<link href="/serviciolocal/shared/css/tables_gray.css" rel="StyleSheet" type="text/css">-->
<title><%= Session("Titulo")%> Contracts - buques</title>
<script src="/serviciolocal/shared/js/fn_windows.js"></script>
<script src="/serviciolocal/shared/js/fn_confirm.js"></script>
<script src="/serviciolocal/shared/js/fn_ayuda.js"></script>
<script src="/serviciolocal/shared/js/fn_fechas.js"></script>
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


function Buscar(){
	var tieneotro;
	var estado;
	document.datos.filtro.value = "";
	tieneotro = "no";
	estado = "si";
/*
	// fec. desde
	if (document.datos.fechadesde.value != ""){
		if (tieneotro == "si"){
			document.datos.filtro.value += " AND " ;
			tieneotro = "si";
		}
		if (validarfecha(document.datos.fechadesde)){
			document.datos.filtro.value += " for_contract.confec >= " + cambiafecha(document.datos.fechadesde.value,true,1) + "";
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
			document.datos.filtro.value += " for_contract.confec <= " + cambiafecha(document.datos.fechahasta.value,true,1) + "";
		}else{
			estado = "no";
		}
	}
	if (!menorque(document.datos.fechadesde.value ,document.datos.fechahasta.value)){
		alert('La fecha hasta debe ser mayor que la fecha desde.');
		estado = "no";
	}	
	*/
	// fec. desde
	if (document.datos.fechadesde.value != ""){
		if (tieneotro == "si"){
			document.datos.filtro.value += " AND " ;
			tieneotro = "si";
		}
		if (validarfecha(document.datos.fechadesde)){
			document.datos.filtro.value += " buq_buque.buqfechas >= " + cambiafecha(document.datos.fechadesde.value,true,1) + "";
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
			document.datos.filtro.value += " buq_buque.buqfechas <= " + cambiafecha(document.datos.fechahasta.value,true,1) + "";
		}else{
			estado = "no";
		}
	}
	if (!menorque(document.datos.fechadesde.value ,document.datos.fechahasta.value)){
		alert('La fecha hasta debe ser mayor que la fecha desde.');
		estado = "no";
	}	
	
	// Tipo Operacion
	if (document.datos.tipopenro.value != 0){
		if (tieneotro == "si"){
			document.datos.filtro.value += " AND buq_buque.tipopenro = '" + document.datos.tipopenro.value + "'";
		}else{
			document.datos.filtro.value += " buq_buque.tipopenro = '" + document.datos.tipopenro.value + "'";
		}
		tieneotro = "si";
	}	
	// Tipo Buque
	if (document.datos.tipbuqnro.value != 0){
		if (tieneotro == "si"){
			document.datos.filtro.value += " AND buq_buque.tipbuqnro = '" + document.datos.tipbuqnro.value + "'";
		}else{
			document.datos.filtro.value += " buq_buque.tipbuqnro = '" + document.datos.tipbuqnro.value + "'";
		}
		tieneotro = "si";
	}		
	// Agencia
	if (document.datos.agenro.value != 0){
		if (tieneotro == "si"){
			document.datos.filtro.value += " AND buq_buque.agenro = '" + document.datos.agenro.value + "'";
		}else{
			document.datos.filtro.value += " buq_buque.agenro = '" + document.datos.agenro.value + "'";
		}
		tieneotro = "si";
	}			
	
	/*
	// Term
	if (document.datos.ternro.value != 0){
		document.datos.filtro.value += " for_contract.ternro = " + document.datos.ternro.value;
		tieneotro = "si";
	}

	// Company
	if (document.datos.comnro.value != 0){
		if (tieneotro == "si"){
			document.datos.filtro.value += " AND for_contract.comnro = " + document.datos.comnro.value;
		}else{
			document.datos.filtro.value += " for_contract.comnro = " + document.datos.comnro.value;
		}
		tieneotro = "si";
	}
	
	// Client
	if (document.datos.clinro.value != 0){
		if (tieneotro == "si"){
			document.datos.filtro.value += " AND for_contract.clinro = " + document.datos.clinro.value;
		}else{
			document.datos.filtro.value += " for_contract.clinro = " + document.datos.clinro.value;
		}
		tieneotro = "si";
	}
	
	// Port
	if (document.datos.pornro.value != 0){
		if (tieneotro == "si"){
			document.datos.filtro.value += " AND for_contract.pornro = " + document.datos.pornro.value;
		}else{
			document.datos.filtro.value += " for_contract.pornro = " + document.datos.pornro.value;
		}
		tieneotro = "si";
	}	
	
	// Product
	if (document.datos.pronro.value != 0){
		if (tieneotro == "si"){
			document.datos.filtro.value += " AND for_contract.pronro = " + document.datos.pronro.value;
		}else{
			document.datos.filtro.value += " for_contract.pronro = " + document.datos.pronro.value;
		}
		tieneotro = "si";
	}		
	
	// ctrnum
	if (document.datos.ctrnum.value != 0){
		if (tieneotro == "si"){
			document.datos.filtro.value += " AND " ;
			tieneotro = "si";
		}
		switch  (parseInt(document.datos.ctrnum.value)){
			case 1: // Comienza con
				document.datos.filtro.value += " ctrnum LIKE '" + document.datos.txtctrnum.value + escape("%") + "'" ;
				break;
			case 2: // Cotiene
				document.datos.filtro.value += " ctrnum LIKE '" + escape("%") + document.datos.txtctrnum.value + escape("%") + "'";
				break;
			case 3: // Igual a
				document.datos.filtro.value += " ctrnum = '" + document.datos.txtctrnum.value + "'";
				break;
		}
		tieneotro = "si";
	}
	*/
	if (estado == "si"){
		window.ifrm.location = 'buques_con_01.asp?asistente=0&filtro=' + document.datos.filtro.value;
	}
}


function Nuevo_Dialogo(w_in, pagina, ancho, alto)
{
 return w_in.showModalDialog(pagina,'', 'center:yes;dialogWidth:' + ancho.toString() + ';dialogHeight:' + alto.toString() + ';');
}
function Ayuda_Fecha(txt)
{
 var jsFecha = Nuevo_Dialogo(window, '/serviciolocal/shared/js/calendar.html', 16, 15);

 if (jsFecha == null) txt.value = ''
 else txt.value = jsFecha;
}

function Limpiar(){

	document.datos.fechadesde.value = "<%= date() - 1 %>";
	document.datos.fechahasta.value = "<%= date() %>";

	document.datos.tipopenro.value     = 0;
	document.datos.tipbuqnro.value     = 0;
	document.datos.agenro.value        = 0;	
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
<body leftmargin="0" topmargin="0" rightmargin="0" bottommargin="0" onload="Javascript:document.datos.tipbuqnro.focus();">
      <table border="0" cellpadding="0" cellspacing="0" height="100%" width="100%">
        <tr style="border-color :CadetBlue;">
          <td align="left" class="barra">&nbsp;</td>
          <td nowrap align="right" class="barra">
		  		  
          <% call MostrarBoton ("opcionbtn", "Javascript:abrirVentana('buques_con_02.asp?Tipo=A','',490,300);","Alta")%>
          <% call MostrarBoton ("opcionbtn", "Javascript:eliminarRegistro(document.ifrm,'buques_con_04.asp?cabnro=' + document.ifrm.datos.cabnro.value);","Baja")%>
          <% call MostrarBoton ("opcionbtn", "Javascript:abrirVentanaVerif('buques_con_02.asp?Tipo=M&cabnro=' + document.ifrm.datos.cabnro.value,'',490,300);","Modifica")%>
		  &nbsp;&nbsp;
          <%' call MostrarBoton ("sidebtnSHW", "Javascript:llamadaexcel();","Excel")%>
		  &nbsp;&nbsp;
          <% call MostrarBoton ("opcionbtn", "Javascript:Contenido();","Contenido")%>		  
		  		  
          <%' call MostrarBoton ("opcionbtn", "Javascript:orden('../../config/contracts_con_01.asp','',490,300);","Orden")%>		  
		  <!--
		  <a class=sidebtn href="Javascript:orden('../../config/contracts_con_01.asp');">Orden</a>
		  -->
		  &nbsp;&nbsp;
		  
		  <% call MostrarBoton ("opcionbtn", "Javascript:abrirVentana('rep_exp_buques_con_00.asp','',780,560);","Estadísticas")%>						  
		  &nbsp;&nbsp;&nbsp;
		  <% call MostrarBoton ("opcionbtn", "Javascript:abrirVentana('rep_exp_buques_x_exp_con_00.asp','',780,560);","Buques/Exportadora")%>	  
		  
		  &nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;		  
		  
		  <% call MostrarBoton ("opcionbtn", "Javascript:abrirVentana('../ess/index.asp','',780,560);","Web")%>		  
		  &nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;		  
		  </td>
        </tr>		
		<tr>
			<td align="center" colspan="2">
				<table border="0">
					<form name="datos">
					<input type="hidden" name="filtro" value="">
					<!--
					<tr>
						<td align="right"><b>Fec. Desde: </b></td>
						<td><input  type="text" name="fechadesde" size="10" maxlength="10" value="<%'= Date() - 1 %>" >
							<a href="Javascript:Ayuda_Fecha(document.datos.fechadesde);"><img src="/serviciolocal/shared/images/cal.gif" border="0"></a>
						</td>
						<td align="right"><b>Fec. Hasta: </b></td>
						<td><input  type="text" name="fechahasta" size="10" maxlength="10" value="<%'= Date() %>" >
							<a href="Javascript:Ayuda_Fecha(document.datos.fechahasta);"><img src="/serviciolocal/shared/images/cal.gif" border="0"></a>
						</td>
					</tr>
					-->
					<tr>
						<td align="right"><b>Fec. Desde: </b></td>
						<td><input  type="text" name="fechadesde" size="10" maxlength="10" value="<%= Date() - 1 %>" >
							<a href="Javascript:Ayuda_Fecha(document.datos.fechadesde);"><img src="/serviciolocal/shared/images/cal.gif" border="0"></a>
						</td>
						<td align="right"><b>Fec. Hasta: </b></td>
						<td><input  type="text" name="fechahasta" size="10" maxlength="10" value="<%= Date() %>" >
							<a href="Javascript:Ayuda_Fecha(document.datos.fechahasta);"><img src="/serviciolocal/shared/images/cal.gif" border="0"></a>
						</td>
					</tr>
					<tr>
						<td  align="right" nowrap><b>Tipo Operación: </b></td>
						<td><select name="tipopenro" size="1" style="width:150;">
								<option value=0 selected>Todos</option>
								<%Set l_rs = Server.CreateObject("ADODB.RecordSet")
								l_sql = "SELECT  * "
								l_sql  = l_sql  & " FROM buq_tipoope "
								l_sql  = l_sql  & " ORDER BY tipopedes "
								rsOpen l_rs, cn, l_sql, 0
								do until l_rs.eof		%>	
								<option value= <%= l_rs("tipopenro") %> > 
								<%= l_rs("tipopedes") %> (<%=l_rs("tipopenro")%>) </option>
								<%	l_rs.Movenext
								loop
								l_rs.Close %>
							</select>
							<script>document.datos.tipopenro.value= "0"</script>
						</td>					
						<td  align="right" nowrap><b>Tipo Buque: </b></td>
						<td><select name="tipbuqnro" size="1" style="width:150;">
								<option value=0 selected>Todos</option>
								<%Set l_rs = Server.CreateObject("ADODB.RecordSet")
								l_sql = "SELECT  * "
								l_sql  = l_sql  & " FROM buq_tipobuque "
								l_sql  = l_sql  & " ORDER BY tipbuqdes "
								rsOpen l_rs, cn, l_sql, 0
								do until l_rs.eof		%>	
								<option value= <%= l_rs("tipbuqnro") %> > 
								<%= l_rs("tipbuqdes") %> (<%=l_rs("tipbuqnro")%>) </option>
								<%	l_rs.Movenext
								loop
								l_rs.Close %>
							</select>
							<script>document.datos.tipbuqnro.value= "0"</script>
						</td>
					</tr>
					<tr>
						<td  align="right" nowrap><b>Agencia: </b></td>
						<td><select name="agenro" size="1" style="width:150;">
								<option value=0 selected>Todas</option>
								<%Set l_rs = Server.CreateObject("ADODB.RecordSet")
								l_sql = "SELECT  * "
								l_sql  = l_sql  & " FROM buq_agencia "
								l_sql  = l_sql  & " ORDER BY agedes "
								rsOpen l_rs, cn, l_sql, 0
								do until l_rs.eof		%>	
								<option value= <%= l_rs("agenro") %> > 
								<%= l_rs("agedes") %> (<%=l_rs("agenro")%>) </option>
								<%	l_rs.Movenext
								loop
								l_rs.Close %>
							</select>
							<script>document.datos.agenro.value= "0"</script>
						</td>					
					</tr>					
					<!--
					<tr>
						<td  align="right"><b>Product: </b></td>
						<td><select name="pronro" size="1" style="width:210;">
								<option value=0 selected>Todos</option>
								<% 'Set l_rs = Server.CreateObject("ADODB.RecordSet")
								'l_sql = "SELECT pronro, prodesabr "
								'l_sql  = l_sql  & " FROM for_product "
								'l_sql  = l_sql  & " ORDER BY prodesabr"
								'rsOpen l_rs, cn, l_sql, 0
								'do until l_rs.eof		%>	
								<option value= <%'= l_rs("pronro") %> > 
								<%'= l_rs("prodesabr") %> (<%'=l_rs("pronro")%>) </option>
								<%	'l_rs.Movenext
								'loop
								'l_rs.Close %>
							</select>
							<script> document.datos.pronro.value= "0"</script>
						</td>						
						<td  align="right"><b>Term: </b></td>
						<td><select name="ternro" size="1" style="width:210;">
								<option value=0 selected>Todos</option>
								<%'Set l_rs = Server.CreateObject("ADODB.RecordSet")
								'l_sql = "SELECT ternro, terdes"
								'l_sql  = l_sql  & " FROM for_term "
								'l_sql  = l_sql  & " ORDER BY terdes"
								'rsOpen l_rs, cn, l_sql, 0
								'do until l_rs.eof		%>	
								<option value= <%'= l_rs("ternro") %> > 
								<%'= l_rs("terdes") %> (<%'=l_rs("ternro")%>) </option>
								<%'	l_rs.Movenext
								'loop
								'l_rs.Close %>
							</select>
							<script> document.datos.ternro.value= "0"</script>
						</td>
					</tr>						
				    -->

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

						<td align="right" colspan="2">
							<table border="0" cellpadding="0" cellspacing="0" bgcolor="Navy">
								<tr>
								
										<!--<td ><img src="../shared/images/gen_rep/boton_01.gif" width="5.9"></td>-->
										<td ><a class="sidebtnABM" href="Javascript:Buscar();">Filtrar</a></td>
										<!--<td  background="../shared/images/gen_rep/boton_05.gif"><img src="../shared/images/gen_rep/boton_03.gif" height="15"></td>-->
										<td ><a class="sidebtnABM" href="Javascript:Limpiar();">Limpiar</a></td>
										<!--<td ><img src="../shared/images/gen_rep/boton_06.gif"></td>-->
								</tr>
							</table>
						</td>
					</tr>					

				</table>
			</td>
		</tr>		
		
        <tr valign="top" height="100%">
          <td colspan="2" style="" width="100%">
      	  <iframe scrolling="yes" name="ifrm" src="buques_con_01.asp" width="100%" height="100%"></iframe> 
	      </td>
        </tr>		
			</form>		
      </table>
</body>

<script>
	Buscar();
</script>
</html>

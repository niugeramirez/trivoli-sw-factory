<% Option Explicit %>
<!--#include virtual="/ess/ess/shared/inc/const.inc"-->
<!--#include virtual="/serviciolocal/shared/db/conn_db.inc"-->
<!--#include virtual="/serviciolocal/shared/inc/fecha.inc"-->
<!--#include virtual="/serviciolocal/ess/shared/inc/accesoESS.inc"-->

<%

'Archivo: novedades_empleado_liq_02.asp
'Descripción: abm de novedades empleado
'Autor: Fernando Favre
'Fecha: 14-10-03
'Modificado: 
'	15-10-03 FFavre Se agrego segundo apellido y nombre
'	17-11-03 FFavre Se agrego firma.
'	17-11-03 FFavre Se agrego la utilizacion de la tecla tab
'	19-11-03 FFavre Se alinearon los campos conccod, tpanro, nevalor
'	19-11-03 FFavre Se agrego la tecla tab con la misma funcion que la tecla enter.
'	24-11-03 FFavre Se agregaron periodos retroactivos.
'	04-02-04 FFavre Se validan las fechas.
'					Se chequean los períodos (Desde no sea menor que Hasta).
'					Se cambio el orden en que se muestran los períodos.
'					En el valor se permiten la cantidad de decimales definidos para el concepto.
'	12-02-04 FFavre Al blanquear los controles, se deshabilitan las fechas y periodos.
'   03-09-04 - Scarpa D. - Validacion de los rangos de vigencias	
'   05-10-04 - Scarpa D. - Correccion de las novedades retroactivas	
'   23-11-04 - Alvaro Bayon - Corrección de búsqueda conceptos y parámetros
'	25-10-05 - Leticia A. - Adecuacion a Autogestion - se comento lo reacionado con Firmas
'	05-12-2005 - Leticia A. - Agregar un campo de texto libre: Motivo para Mega.

 on error goto 0
 
 Dim l_sql
 Dim l_rs
 Dim l_tipo
 
 Dim l_ternro
 Dim l_concnro
 Dim l_conccod
 Dim l_concabr
 Dim l_tpanro
 Dim l_tpadabr
 Dim l_nenro
 Dim l_nevalor
 Dim l_unisigla
 Dim l_nevigencia
 Dim l_nedesde
 Dim l_nehasta
 Dim l_concretro
 Dim l_nepliqdesde
 Dim l_nepliqhasta
 Dim l_netexto
 Dim l_pronro
 Dim l_empleg
 Dim l_ApNombre
 Dim l_retro
 Dim l_selectdesde
 Dim l_selecthasta
 Dim l_conccantdec
 
 l_ternro  = l_ess_ternro 
 l_concnro = request("concnro")
 l_tpanro  = request("tpanro")
 l_tipo	   = request("tipo")
 l_retro   = request("retro")
 l_nenro   = request("nenro") 
 
 Set l_rs = Server.CreateObject("ADODB.RecordSet")
 
'--------------------------------------------------------------------------------------------------------------------
' Se buscan los datos del empleado
 if l_ternro <> "" then
	l_sql = "SELECT empleg, ternro, terape, ternom, terape2, ternom2 FROM empleado WHERE ternro = " & l_ternro
	rsOpen l_rs, cn, l_sql, 0
	if not l_rs.eof then
	
		l_empleg   = l_rs("empleg")
		l_ApNombre = l_rs("terape") 
		if l_rs("terape2") <> "" then
			l_ApNombre = l_ApNombre & " " & l_rs("terape2")
		end if
		l_ApNombre = l_ApNombre & " " & l_rs("ternom")
		if l_rs("ternom2") <> "" then
			l_ApNombre = l_ApNombre & " " & l_rs("ternom2")
		end if
	end if
	l_rs.Close
 end if
 
 if l_retro = "" then
 	l_retro = false
 end if
 
'--------------------------------------------------------------------------------------------------------------------
' Se cargan los datos, dependiendo si es un alta o modificacion
 select Case l_tipo
	Case "A":
		l_concnro     = 0
		l_tpanro      = null
		l_nevalor     = null
		l_nevigencia  = 0
		l_nedesde     = null
		l_nehasta     = null
		l_nepliqdesde = 0
		l_nepliqhasta = 0
		l_netexto = ""
		l_pronro      = null
		l_nenro       = 0
	Case "M":
		l_sql = "SELECT novemp.nenro, concepto.conccod, concepto.concretro, concepto.concabr, concepto.conccantdec, "
		l_sql = l_sql & "tipopar.tpadabr, novemp.concnro, novemp.tpanro, novemp.nevalor, novemp.nevigencia, "
		l_sql = l_sql & "novemp.nedesde, novemp.nehasta, novemp.nepliqdesde, novemp.nepliqhasta, novemp.netexto, unidad.unisigla "
		l_sql = l_sql & "FROM novemp INNER JOIN concepto ON novemp.concnro = concepto.concnro "
		l_sql = l_sql & "INNER JOIN tipopar ON novemp.tpanro = tipopar.tpanro "
		l_sql = l_sql & "INNER JOIN unidad ON tipopar.uninro = unidad.uninro "
		l_sql = l_sql & "WHERE novemp.nenro =" & l_nenro
		rsOpen l_rs, cn, l_sql, 0
		if not l_rs.eof then
			l_conccod 	  = l_rs("conccod")
			l_concabr	  = l_rs("concabr")
			l_tpadabr 	  = l_rs("tpadabr")
			l_nevalor     = replace(l_rs("nevalor"), ",", ".")
			l_unisigla    = l_rs("unisigla")
			l_nevigencia  = l_rs("nevigencia")
			l_nedesde     = l_rs("nedesde")
			l_nehasta     = l_rs("nehasta")
			l_concretro   = l_rs("concretro")
			l_conccantdec = l_rs("conccantdec")
			if l_rs("nepliqdesde") <> "" then
				l_nepliqdesde = l_rs("nepliqdesde")
			else
				l_nepliqdesde = 0
			end if
			if l_rs("nepliqhasta") <> "" then
				l_nepliqhasta = l_rs("nepliqhasta")
			else
				l_nepliqhasta = 0
			end if
			if not isNull(l_rs("netexto")) then
				l_netexto = l_rs("netexto")
			else
				l_netexto= ""
			end if
		end if
		l_rs.Close
 end select
 
'--------------------------------------------------------------------------------------------------------------------
' Select de períodos desde y de períodos hasta
 l_sql = "SELECT pliqnro, pliqdesc, pliqanio, pliqmes "
 l_sql = l_sql & "FROM periodo "
 l_sql = l_sql & "ORDER BY pliqanio DESC, pliqmes DESC"
 rsOpen l_rs, cn, l_sql, 0 
 l_selectdesde = l_selectdesde & "<option value='0,,' selected>Ninguno</option>"
 l_selecthasta = l_selecthasta & "<option value='0,,' selected>Ninguno</option>"
 do until l_rs.eof
 	if l_rs("pliqnro") = l_nepliqdesde then
		l_selectdesde = l_selectdesde & "<option selected value=" & l_rs("pliqnro") & "," & l_rs("pliqanio") & "," & l_rs("pliqmes") & ">" & l_rs("pliqdesc") & "</option>"
	else
		l_selectdesde = l_selectdesde & "<option value=" & l_rs("pliqnro") & "," & l_rs("pliqanio") & "," & l_rs("pliqmes") & ">" & l_rs("pliqdesc") & "</option>"
	end if
 	if l_rs("pliqnro") = l_nepliqhasta then
		l_selecthasta = l_selecthasta & "<option selected value=" & l_rs("pliqnro") & "," & l_rs("pliqanio") & "," & l_rs("pliqmes") & ">" & l_rs("pliqdesc") & "</option>"
	else
		l_selecthasta = l_selecthasta & "<option value=" & l_rs("pliqnro") & "," & l_rs("pliqanio") & "," & l_rs("pliqmes") & ">" & l_rs("pliqdesc") & "</option>"
	end if
	l_rs.movenext
 loop
 l_rs.close
 
'--------------------------------------------------------------------------------------------------------------------
' Firmas
' Dim l_tipAutorizacion  'Es el tipo del circuito de firmas
' Dim l_HayAutorizacion  'Es para ver si las autorizaciones estan activas
' Dim l_PuedeVer         'Es para ver si las autorizaciones estan activas
 
 'l_sql = "select cystipo.* from cystipo "
 'l_sql = l_sql & "where (cystipo.cystipact = -1) and cystipo.cystipnro = 5 "
 'rsOpen l_rs, cn, l_sql, 0 
 
 'l_HayAutorizacion = not l_rs.eof
 'if not l_rs.eof then
 	'l_tipAutorizacion = 5
 'end if 
 'l_rs.close
 
 'if l_HayAutorizacion AND (l_tipo = "M") then
	'l_sql = "select cysfirautoriza, cysfirsecuencia, cysfirdestino from cysfirmas "
	'l_sql = l_sql & "where cysfirmas.cystipnro = " & l_tipAutorizacion & " and cysfirmas.cysfircodext = '" & l_nenro & "' " 
	'l_sql = l_sql & "order by cysfirsecuencia desc"
	'rsOpen l_rs, cn, l_sql, 0 
 	
	'l_PuedeVer = False
 	
	'if not l_rs.eof then
 		'if (l_rs("cysfirautoriza") = session("UserName")) or (l_rs("cysfirdestino") = session("UserName")) then 
	   		'Es una modificación del ultimo o es el nuevo que autoriza 
    		'l_PuedeVer = True 
    	'end if
 	'end if
	'l_rs.close
	
 	'If not l_PuedeVer then
    	'response.write "<script>alert('No esta autorizado a ver o modificar este registro.');window.close()</script>"
		'response.end
 	'End if
 'End if
%>
<html>
<head>
<link href="../<%=c_estilo %>" rel="StyleSheet" type="text/css">
<meta http-equiv="Content-Type" content="text/html; charset=iso-8859-1">
<title>Novedades por Empleado - Liquidaci&oacute;n de Haberes - RHPro &reg;</title>
</head>
<script src="/serviciolocal/shared/js/fn_windows.js"></script>
<script src="/serviciolocal/shared/js/fn_ayuda.js"></script>
<script src="/serviciolocal/shared/js/fn_ay_generica.js"></script>
<script src="/serviciolocal/shared/js/fn_numeros.js"></script>
<script src="/serviciolocal/shared/js/fn_fechas.js"></script>
<script src="/serviciolocal/shared/js/fn_valida.js"></script>
<script>

function Validar_Formulario(){
 var par;
 var seguir;
	seguir = '0';
	if (document.all.conccod.value == ""){
		alert("Primero debe seleccionar un concepto.");
		document.all.conccod.focus();
		document.all.conccod.select();
	}
	else
 	if (document.all.tpanro.value == ""){
		alert("Primero debe seleccionar un parámetro.");
		document.all.tpanro.focus();
		document.all.tpanro.select();
	}
	else
 	if (document.all.nevalor.value == ""){
		alert("El valor no puede ser nulo.");
		document.all.nevalor.focus();
		document.all.nevalor.select();
	}
	else
	if (!validanumero(document.all.nevalor, 15, document.all.conccantdec.value)){
		alert("El valor no es válido. Se permite hasta 15 enteros y "+document.all.conccantdec.value+" decimales.\nEl concepto tiene "+document.all.conccantdec.value+" decimales definidos.");
		document.all.nevalor.focus();
		document.all.nevalor.select();
	}
	else
	if (document.all.nevigencia.checked && !validarfecha(document.all.nedesde)){
		document.all.nedesde.focus();
		document.all.nedesde.select();
	}
	else
	if (document.all.nevigencia.checked && (Trim(document.all.nehasta.value) != '') && !validarfecha(document.all.nehasta)){
		document.all.nehasta.focus();
		document.all.nehasta.select();
	}
	else
	if (document.all.nevigencia.checked && (Trim(document.all.nehasta.value) != '') && !menorque(document.all.nedesde.value, document.all.nehasta.value)){
		alert('La fecha desde es mayor que la fecha hasta.');
		document.all.nedesde.focus();
		document.all.nedesde.select();
	}
	else
	if (document.all.concretro.checked)
	{
		var valperdesde = document.all.nepliqdesde.value.split(",");
		var valperhasta = document.all.nepliqhasta.value.split(",");
		var pliqnrodesde = valperdesde[0];
		var anioperdesde = valperdesde[1];
		var mesperdesde  = valperdesde[2];
		var pliqnrohasta = valperhasta[0];
		var anioperhasta = valperhasta[1];
		var mesperhasta  = valperhasta[2];
		if (pliqnrodesde == 0){
			alert('Debe seleccionar un Período Desde.');
			document.all.nepliqdesde.focus();
		}
		else
		if (pliqnrohasta == 0){
			alert('Debe seleccionar un Período Hasta.');
			document.all.nepliqhasta.focus();
		}
		else
		if ((parseInt(anioperdesde) > parseInt(anioperhasta) && anioperhasta != '') || ((parseInt(anioperdesde) == parseInt(anioperhasta)) && (parseInt(mesperdesde) > parseInt(mesperhasta)))){
			alert('El Período Desde (' + mesperdesde + '/' + anioperdesde + ') es mayor que el Período Hasta (' + mesperhasta + '/' + anioperhasta + ')');
			document.all.nepliqdesde.focus();
		}
		else
			seguir = '-1';
	}
	else
		seguir = '-1';

	if (seguir == '-1')
	<%' if l_HayAutorizacion then ' Si se debe tomar autorizacion %>
		// Verifico que se haya cargado la autorización 
		<%'if l_tipo = "A" then%>
		//if ((document.datos.seleccion.value == "") && (document.datos.seleccion1.value == ""))
		  //  alert("Debe ingresar una autorización.");
		//else{
		<%' else %>
			//if ( '<%'=not l_PuedeVer%>' == 'True')
				//alert('No esta autorizado a ver o modificar este registro.')
			//else{
		<%'end if%>
	<% 'else%>
		{
	<% 'End If %>
		par = "tipo=<%=  l_tipo%>"
		par = par + "&empleado=" + document.all.ternro.value; 
		par = par + "&concnro=" + document.all.concnro.value;
		par = par + "&tpanro=" + document.all.tpanro.value;
		par = par + "&nevalor=" + document.all.nevalor.value * 10000;
		par = par + "&nedesde=" + document.all.nedesde.value;
		par = par + "&nehasta=" + document.all.nehasta.value;				
		par = par + "&nenro="  + document.all.nenro.value;						
		document.valida.location = "novedades_empleado_liq_06.asp?" + par;
	}
}

function Valido(){

 	abrirVentanaH('','bblank',0,0);	
	document.datos.submit();
}

function invalido(causa){
	switch (causa){
		case 'duplicada':
			alert('Ya existe novedad para el concepto y parámetro.');
			Blanquear();
			break;
		case 'ConcNoExiste':
			alert('El concepto no existe.');
			break;
		case 'NoConceptoInd':
			alert('El concepto no se encuentra configurado como individual.');
			break;
		case 'ParNoValido':
			alert('El Parámetro no pertenece al Concepto.');
			break;
		case 'Vigencia01':
			alert('Existe una novedad sin vigencia para el concepto y parámetro.');
			break;
		case 'Vigencia02':
			alert('La vigencia de la novedad se superpone con otra.');
			break;
		case 'Vigencia03':
			alert('Ya existe otra novedad con la fecha hasta de la vigencia vacía.');
			break;
	}
}

function Ayuda_Fecha(txt){
	var jsFecha = Nuevo_Dialogo(window, '/serviciolocal/shared/js/calendar.html', 16, 15);
 	if (jsFecha != null) 
		txt.value = jsFecha;
}

function TeclaTab(donde){
	switch (donde){
		case 'c':
			if (stringValido(document.all.conccod.value)){
				par = "tipo=concepto";
				par = par + "&conccod=" + document.all.conccod.value;
				document.valida.location = "novedades_empleado_liq_06.asp?" + par;
				}
			else{
				alert("El Concepto contiene caracteres no válidos");
				document.all.conccod.focus()
				}
			break;
		case 'p':
			if (stringValido(document.all.conccod.value))
				if (validanumero(document.all.tpanro,9,0)){
					if (document.all.tpanrold.value != document.all.tpanro.value)
						par = "tipo=parametro";
					
					par = par + "&concnro="+document.all.concnro.value;
					par = par + "&tpanro="+document.all.tpanro.value;
					document.validaPar.location = "novedades_empleado_liq_06.asp?" + par;
					}
				else{
					alert("El Parámetro debe ser un número sin decimales");
					document.all.tpanro.focus()
					}
			else{
				alert("El Concepto contiene caracteres no válidos");
				document.all.conccod.focus()
				}
			break;
	}
}

function Tecla(num, donde){
	if (num==13) {
		switch (donde){
			case 'c':
				document.all.conccod.blur();
				break;
			case 'p':
				document.all.tpanro.blur();
				break;
			case 'v':
				Validar_Formulario();
		}
		return false;
	}
	return num;
}

function Actualizar() {
	if (document.all.nevigencia.checked){
		document.all.nedesde.className = "habinp";
		document.all.nehasta.className = "habinp";
		document.all.nedesde.tabIndex = 5;
		document.all.nehasta.tabIndex = 6;
		document.all.nedesde.disabled = false;
		document.all.nehasta.disabled = false;
	}
	else { 
		document.all.nedesde.value = "";
		document.all.nehasta.value = "";
		document.all.nedesde.className = "deshabinp";
		document.all.nehasta.className = "deshabinp";
		document.all.nedesde.tabIndex = -1;
		document.all.nehasta.tabIndex = -1;
		document.all.nedesde.disabled = true;
		document.all.nehasta.disabled = true;
	}
}

function Blanquear(){
	document.all.concnro.value = "";
	document.all.conccod.value = "";
	document.all.concabr.value = "";
	document.all.tpanro.value  = "";
	document.all.tpadabr.value = "";
	document.all.nevalor.value = "";
	document.all.unidesc.value = "";
	document.all.nedesde.value = "";
	document.all.nehasta.value = "";
	document.all.nedesde.disabled = true;
	document.all.nehasta.disabled = true;
	document.all.nedesde.className = "deshabinp";
	document.all.nehasta.className = "deshabinp";
	document.all.nedesde.tabIndex = -1;
	document.all.nehasta.tabIndex = -1;
	document.all.nepliqdesde.value = "0,,";
	document.all.nepliqhasta.value = "0,,";
	document.all.nepliqdesde.disabled = true;
	document.all.nepliqhasta.disabled = true;
	document.all.nepliqdesde.className = "deshabinp";
	document.all.nepliqhasta.className = "deshabinp";
	document.all.nepliqdesde.tabIndex = -1;
	document.all.nepliqhasta.tabIndex = -1;
	document.all.nevigencia.checked  = false;
	document.all.concretro.disabled = false;
	document.all.concretro.checked  = false;
	document.all.concretro.disabled = true;
	document.all.conccod.focus();
}

function AlCargar(){
	if ("<%= l_tipo%>"=="A"){
		document.all.conccod.focus();
		document.all.conccod.select();
	}
	else{
		document.all.nevalor.focus();
		document.all.nevalor.select();
	}
}

function Salir(){
	window.opener.ifrm.location.reload();
	window.close();
}

function Conceptos(){
<%	if l_tipo = "A" then%>
		abrirVentana('help_conceptos_liq_00.asp','',650,280)
<% 	else %>
		alert('No se puede modificar el concepto.');
<%	end if %>		
}

function Parametros(){
	if (document.all.conccod.value == "")
		alert("Primero debe seleccionar un concepto.");
	else
	if (document.all.concnro.value == 0)
		document.valida.location = "novedades_empleado_liq_06.asp?tipo=concepto&conccod=" + document.all.conccod.value;
	else
	<%	if l_tipo = "A" then%>
		abrirVentana('help_parametros_liq_00.asp?concnro='+document.all.concnro.value,'',350,200)
	<% 	else %>
		alert('No se puede modificar el parámetro.');
	<%	end if %>		
}

function AsignarPeriodo(nepliq, pliq){
	var valper = pliq.value.split(",");
	if (valper[0] == 0)
		nepliq.value = "";
	else
		nepliq.value = valper[0];
}

function Firmas(){  // Para llamar a control de firmas, mandandole la descripcion y demas
  // abrirVentana('../gti/cysfirmas_00.asp?obj=document.all.seleccion&amp;tipo=<%'= l_tipAutorizacion %>&amp;codigo=<%'= l_nenro %>&amp;descripcion=Nov Liq','_blank','toolbar=no,location=no,directories=no,status=no,menubar=no,scrollbars=no,resizable=yes,width=421,height=180')
}

function cambioRetro(){
  if (!document.datos.concretro.disabled){
	  if (document.datos.concretro.checked){
	     document.datos.nepliqdesde.disabled = false;
	     document.datos.nepliqhasta.disabled = false;	 
	     document.datos.nepliqdesde.className = 'habinp';
	     document.datos.nepliqhasta.className = 'habinp';
	  }else{
	     document.datos.nepliqdesde.disabled = true;  
	     document.datos.nepliqhasta.disabled = true;	 	 
	     document.datos.nepliqdesde.className = 'deshabinp';
	     document.datos.nepliqhasta.className = 'deshabinp';	 
	  }
  }else{
	     document.datos.nepliqdesde.disabled = true;  
	     document.datos.nepliqhasta.disabled = true;	 	 
	     document.datos.nepliqdesde.className = 'deshabinp';
	     document.datos.nepliqhasta.className = 'deshabinp';	 
  }
}

function actualizarConcRetro(tipo){
  if (tipo == 0){
     document.datos.concretro.disabled = true;
  }else{
     document.datos.concretro.disabled = false;
  }
  document.datos.concretro.checked = false;
  cambioRetro();
}

</script>
<body leftmargin="0" rightmargin="0" topmargin="0" bottommargin="0" onload="Javascript:AlCargar();"> <!--   onunload="Javascript:window.opener.ifrm.location.reload()" -->
<form name="datos" target="bblank" action="novedades_empleado_liq_03.asp?tipo=<%=l_tipo%>&empleg=<%=l_empleg%>" method="post">
<input type="Hidden" name="ternro" value="<%=l_ternro%>">
<input type="Hidden" name="nenro" value="<%=l_nenro%>">
<input type="Hidden" name="concnro" value="<%=l_concnro%>">
<input type="Hidden" name="tpanrold">
<input type="hidden" name="seleccion" value="">
<input type="hidden" name="seleccion1" value="">
<input type="hidden" name="conccantdec" value="<%= l_conccantdec %>">
<input type="hidden" name="nepliqdes" value="<%= l_nepliqdesde %>">
<input type="hidden" name="nepliqhas" value="<%= l_nepliqhasta %>">
<table cellspacing="0" cellpadding="0" border="0" width="100%" height="100%">
	<tr style="border-color :CadetBlue;">
	    <th colspan="4" align="left" class="th2">Datos de la Novedad</th>
		<th colspan="2" align="right"> &nbsp;
		<%' if l_HayAutorizacion then ' Si se debe tomar autorizacion %>
			<% 'call MostrarBoton ("sidebtnSHW", "Javascript:Firmas();","Autorizar") %>
			&nbsp;&nbsp;&nbsp;
		<%' End If %>		
		</th>
	</tr>
	<tr>
		<td align="center" colspan="6">
			<b>Empleado: </b>
			<input type="Text" class="deshabinp" tabindex="-1" readonly size="7" name="empleg" value="<%= l_empleg%>">
			<input type="Text" class="deshabinp" tabindex="-1" readonly size="40" name="ApNombre" value="<%= l_ApNombre%>">
		</td>	
	</tr>
	<tr>
		<td colspan="6">&nbsp;</td>
	</tr>
	<tr>
		<td align="right"><b>Concepto:</b>&nbsp;</td>
		<td>
			<input type="Text" size="12" maxlength="12" name="conccod" value="<%=l_conccod%>" onchange="TeclaTab('c');" onkeypress="return Tecla(event.keyCode,'c')" <%if l_tipo = "M" then%>readonly class="deshabinp" tabindex="-1"<%else%> class="habinp" tabindex="1"<%end if%>>
			&nbsp;
		</td>
		<td colspan="3" align="left">
			<input type="Text" size="35" name="concabr" readonly class="deshabinp" tabindex="-1" value="<%= l_concabr%>">
		</td>
		<td>
			<a class=sidebtnSHW href="Javascript:Conceptos();" style="width: 70px;" tabindex="-1">Conceptos</a>
		</td>
	</tr>
	<tr>
		<td align="right"><b>Par&aacute;metro:</b>&nbsp;</td>
		<td>
			<input type="Text" size="12" maxlength="9" name="tpanro" onchange="TeclaTab('p');" onkeypress="return Tecla(event.keyCode,'p')" value="<%= l_tpanro%>" <%if l_tipo = "M" then%>readonly class="deshabinp" tabindex="-1"<%else%> class="habinp" tabindex="2"<%end if%>>
			&nbsp;
		</td>
		<td colspan="3" align="left">
			<input type="Text" size="35" name="tpadabr" readonly class="deshabinp" tabindex="-1" value="<%= l_tpadabr%>">
		</td>
		<td>
			<a class=sidebtnSHW href="Javascript:Parametros();" style="width: 70px;" tabindex="-1">Par&aacute;metros</a>
		</td>
	</tr>
	<tr>
		<td align="right"><b>Valor:</b> &nbsp;</td>
		<td align="left"><input type="text" size="21" maxlength="20" name="nevalor" tabindex="3" onchange="TeclaTab('v');" onkeypress="return Tecla(event.keyCode,'v')" value="<%= l_nevalor%>" class="habinp"></td>
		<td align="right"><b>Unidad: </b></td>
		<td colspan="3" align="left">
			<input type="Text" size="10" name="unidesc" tabindex="-1" readonly class="deshabinp" value="<%= l_unisigla%>">
		</td>
	</tr>
	<tr>
		<td></td>
		<td align="left">
			<input type="Checkbox" name="nevigencia" tabindex="4" <%if l_nevigencia = true then%>checked<%end if%> onclick="Javascript:Actualizar();">
			<b>Vigencia</b>
		</td>
		<td align="right"><b>Desde:</b></td>
		<td>
			<input type="text" name="nedesde" maxlength="10" size="10" value="<%= l_nedesde%>" <%if l_nevigencia = false then%>disabled class="deshabinp" tabindex="-1"<%else%> tabindex="5"<%end if%>>
			<a href="Javascript:if (document.all.nevigencia.checked) Ayuda_Fecha(document.datos.nedesde);" tabindex="-1"><img src="/serviciolocal/shared/images/cal.gif" border="0"></a>	
		</td>
		<td align="right"><b>Hasta:</b></td>
		<td>
			<input type="text" name="nehasta" maxlength="10" size="10" value="<%= l_nehasta%>" <%if l_nevigencia = false then%>disabled class="deshabinp" tabindex="-1"<%else%> tabindex="6"<%end if%>>
			<a href="Javascript:if (document.all.nevigencia.checked) Ayuda_Fecha(document.datos.nehasta);" tabindex="-1"><img src="/serviciolocal/shared/images/cal.gif" border="0"></a>	
		</td>
	</tr>
	<tr>
		<td></td>
		<td align="left">
		    <% if CInt(l_concretro) = -1 then%>
			<input type="Checkbox" name="concretro" tabindex="-1" onclick="javascript:cambioRetro();" <%if l_nepliqdesde <> 0 then%>checked<%end if%>>			
			<% else %>
			<input type="Checkbox" name="concretro" tabindex="-1" onclick="javascript:cambioRetro();" disabled>
			<% end if %>
			<b>Retroactivo</b>
		</td>
		<td align="right"><b>Desde:</b></td>
		<td align="left">
			<select style="width:130px" name="nepliqdesde" size='1' <%if l_concretro = false then%>disabled class="deshabinp" tabindex="-1"<%else%> tabindex="6"<%end if%> 
			onchange="javascript:AsignarPeriodo(document.all.nepliqdes, document.all.nepliqdesde);">
		    <%= l_selectdesde %>
			</select>
		</td>
		<td align="right"><b>Hasta:</b></td>
		<td align="left">
			<select style="width:130px" name="nepliqhasta" size='1' <%if l_concretro = false then%>disabled class="deshabinp" tabindex="-1"<%else%> tabindex="7"<%end if%>
			onchange="javascript:AsignarPeriodo(document.all.nepliqhas, document.all.nepliqhasta);">
		    <%= l_selecthasta %>
			</select>
		</td>
	</tr>
	<tr>
		<td align="right"> <b>Motivo:</b>&nbsp;</td>
		<td colspan="5"> 
			<input type="Text" name="netexto" value="<%= l_netexto%>" size="30" maxlength="30" class="habinp">
		 </td>
	</tr>
	<tr>
	    <td colspan="6" align="right" class="th2">
			<br>
			<a class=sidebtnABM href="Javascript:Validar_Formulario();" tabindex="8">Aceptar</a> &nbsp;
			<a class=sidebtnABM href="Javascript:Salir();" tabindex="9">Cancelar</a> &nbsp;
			<br>&nbsp;
		</td>
	</tr>
</table>
<iframe name="valida" src="blanc.asp" width="0" height="0"></iframe> 
<iframe name="validaPar" src="blanc.asp" width="0" height="0"></iframe> 
</form>
<%
set l_rs = nothing
%>

<script>
  cambioRetro();
</script>

</body>
</html>

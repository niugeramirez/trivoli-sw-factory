<%Option Explicit %>
<!--#include virtual="/ess/ess/shared/inc/const.inc"-->
<!--#include virtual="/serviciolocal/shared/inc/sec.inc"-->
<!--#include virtual="/serviciolocal/shared/inc/const.inc"-->
<!--#include virtual="/serviciolocal/shared/db/conn_db.inc"-->
<!--#include virtual="/serviciolocal/shared/inc/numero.inc"-->
<!--#include virtual="/serviciolocal/shared/inc/sqls.inc"-->
<!--#include virtual="/serviciolocal/ess/shared/inc/accesoESS.inc"-->

<!--
-----------------------------------------------------------------------------
Archivo        : vales_liq_02.asp
Descripcion    : Modulo que se encarga de mostrar los datos de un vale
Creador        : Scarpa D.
Fecha Creacion : 09/01/2004
Modificacion   :
  27/01/2003 - Scarpa D. - Sacar del titulo la palabra configuracion
                           Copiar la fecha de pedido a fecha prevista
						   Poner por default la moneda origen
						   Periodos ordenados por anio y mes
						   Validar que la fecha de pedido pertenezca al periodo
  09-02-2004 - F. Favre - Se pueden realizar altas continuas. 
  10-02-2004 - F. Favre - Se reacomodaron los campos. 
  08/06/2004 - Alvaro Bayon - Conservo el valnro para la modificación
  25-06-2004 - F.Favre - El campo valmonto no funcionaba para numeros muy muchos digitos. 
  Modificado  : 12/09/2006 Raul Chinestra - se agregó Vales en Autogestión     
                28/09/2006 Maximiliano Breglia - se saco v_empleado 
				26/02/2007 - Martin Ferraro - Inicializar el monto de los vales segun lo config en el tipo de vale 
				30/05/2007 - Martin Ferraro - La modificacion de vales respeta las modificadiones de monto fijo
				11/06/2007 - Manuel L - Se cambió src="" por src="blanc.asp" en un iframe
-----------------------------------------------------------------------------
-->
<% 
on error goto 0

 Dim l_tipo
 
' Variables
 Dim l_valnro
 Dim l_empleado
 Dim l_ppagnro
 Dim l_monnro
 Dim l_valmonto
 Dim l_valfecped
 Dim l_valfecprev
 Dim l_pliqnro
 Dim l_valdesc
 Dim l_pliqdto
 Dim l_pronro
 Dim l_tvalenro
 Dim l_tvalenro2
 Dim l_valrevis
 Dim l_mvalores
 
 
'Variables locales
 Dim l_apnom
 Dim l_empleg
 Dim l_ant_leg
 Dim l_sig_leg
 
 Dim l_rs
 Dim l_rs1
 Dim l_sql
 
 dim l_ternro
 
 Dim l_tvalemfijo
 Dim l_tvalemonto
 
 l_ternro = l_ess_ternro
 l_empleg = l_ess_empleg
 
' response.write "ternro " & l_ternro
' response.write "empleg " & l_empleg
 
 l_tipo     	 = Request.QueryString("tipo")
 l_pliqnro  	 = Request.QueryString("pliqnro")
 l_tvalenro 	 = Request.QueryString("tvalenro")
 l_valnro   	 = Request.QueryString("valnro")
 
 l_ppagnro    = Request.QueryString("ppagnro")
 l_monnro     = Request.QueryString("monnro")
 l_valmonto   = Request.QueryString("valmonto")
 l_valfecped  = Request.QueryString("valfecped")
 l_valfecprev = Request.QueryString("valfecprev")
 l_valdesc    = Request.QueryString("valdesc")
 l_pliqdto    = Request.QueryString("pliqdto")
 l_pronro     = Request.QueryString("pronro")
 l_tvalenro2  = Request.QueryString("tvalenro2")
' l_pliqnro    = Request.QueryString("pliqnro")
 l_valrevis   = Request.QueryString("valrevis")
 'l_empleg	  = Request.QueryString("empleg")
 l_mvalores   = Request.QueryString("mvalores")
 
 Set l_rs = Server.CreateObject("ADODB.RecordSet")	
 
'--------------------------------------------------------------------------------------------------------------------
 select Case l_tipo
	Case "A":
 		
		if CInt(l_mvalores) <> -1 then
		 	l_valnro     = 0
			l_empleado   = 0
			l_ppagnro    = 0
			l_monnro     = 0
			l_valmonto   = 0
			l_valfecped  = ""
			l_valfecprev = ""
			l_valdesc    = ""
			l_pliqdto    = l_pliqnro
			l_tvalenro2  = l_tvalenro
			l_pronro     = 0
	        l_valrevis	 = 0
		end if
		
		if l_tvalenro2 = "" then
		   l_tvalenro2 = 0
		end if
		
	Case "M":
		
		l_sql =         " SELECT vales.*,terape, terape2, ternom, ternom2, empleg, tipovale.* "
		l_sql = l_sql & " FROM vales "
		l_sql = l_sql & " INNER JOIN tipovale ON tipovale.tvalenro = vales.tvalenro "
		l_sql = l_sql & " INNER JOIN empleado ON empleado.ternro = vales.empleado "
		l_sql = l_sql & " WHERE valnro= " & l_valnro
		
		rsOpen l_rs, cn, l_sql, 0 
		if not l_rs.eof then
			
			l_empleado   = l_rs("empleado")
			l_ppagnro    = l_rs("ppagnro")
			l_monnro     = l_rs("monnro")
			if isnull(l_monnro) then
				l_monnro = 0
			end if
			l_valmonto   = l_rs("valmonto")
			l_valfecped  = l_rs("valfecped")
			l_valfecprev = l_rs("valfecprev")
			l_valdesc    = l_rs("valdesc")
			l_pliqdto    = l_rs("pliqdto")
			l_pronro     = l_rs("pronro")
			l_tvalenro2  = l_rs("tvalenro")
			l_pliqnro    = l_rs("pliqnro")
			l_valrevis   = l_rs("valrevis")
			
			l_apnom      = l_rs("terape") & " " & l_rs("terape2") & ", " & l_rs("ternom") & " " & l_rs("ternom2")
			l_empleg     = l_rs("empleg")
			
		end if
		l_rs.Close
 end select
 
'--------------------------------------------------------------------------------------------------------------------
' Busco el Empleado
l_sql = "SELECT empleg, ternro, terape, terape2, ternom, ternom2 FROM empleado WHERE empleg = " & l_empleg
 
 l_rs.Maxrecords = 1
 rsOpen l_rs, cn, l_sql, 0
 if not l_rs.eof then
 	l_empleado 	= l_rs("ternro")
	l_empleg 	= l_rs("empleg")
	l_apnom 	= l_rs("terape")
	if l_rs("terape2") <> "" then
		l_apnom = l_apnom & " " & l_rs("terape2")
	end if
	l_apnom = l_apnom & " " & l_rs("ternom")
	if l_rs("ternom2") <> "" then
		l_apnom = l_apnom & " " & l_rs("ternom2")
	end if
 end if
 l_rs.Close 
 
%>

<html>
<head>
<link href="../<%=c_estilo %>" rel="StyleSheet" type="text/css">
<meta http-equiv="Content-Type" content="text/html; charset=iso-8859-1">
<title>Vales</title>
<script src="/serviciolocal/shared/js/fn_windows.js"></script>
<script src="/serviciolocal/shared/js/fn_confirm.js"></script>
<script src="/serviciolocal/shared/js/fn_hora.js"></script>
<script src="/serviciolocal/shared/js/fn_ayuda.js"></script>
<script src="/serviciolocal/shared/js/fn_fechas.js"></script>
<script src="/serviciolocal/shared/js/fn_ay_generica.js"></script>
<script src="/serviciolocal/shared/js/fn_buscar_emp.js"></script>
<script src="/serviciolocal/shared/js/fn_help_emp.js"></script>
<script src="/serviciolocal/shared/js/fn_numeros.js"></script>
<script>

function Tecla(num, donde){
	var aux = new String("");
	if (num==13) {
		switch (donde){
			case 'emp':
				buscarEmpleado();
				break;
			case 'valor':
				Validar_Formulario();
		}
		return false;
	}
	return num;
}

function Ayuda_Fecha(txt){
	var jsFecha = Nuevo_Dialogo(window, '/serviciolocal/shared/js/calendar.html', 16, 15);
 	if (jsFecha != null) 
		txt.value = jsFecha;
}

function Validar_Formulario(){
    var errores = 0;
    document.datos.valmonto2.value = document.datos.valmonto.value.replace(",", ".");
  
    if (!validanumero(document.datos.valmonto2, 15, 4)){
		  alert("El Monto no es válido. Se permite hasta 15 enteros y 4 decimales.");	
		  document.datos.valmonto.focus();
		  document.datos.valmonto.select();
		  return;
    }		
  
	if (Trim(document.datos.valdesc.value) == "" ){
	      alert("Ingrese una descripción.");
		  document.datos.valdesc.focus();
	      return;
	}  
    if (document.datos.tvalenro2.value == ""){
	      alert("Debe selectar un tipo de vale.");
 		  document.datos.tvalenro2.focus();
	      return;
    }  

    if (document.datos.valfecped.value == ""){
	      alert("Debe ingresar una fecha de pedido.");
		  document.datos.valfecped.focus();
		  return;
    }
  
    if (!validarfecha(document.datos.valfecped)){
	  	  document.datos.valfecped.focus();
		  document.datos.valfecped.select();
	      return;
	}

    if (document.datos.valfecprev.value == ""){
	      alert("Debe ingresar una fecha prevista.");
		  document.datos.valfecprev.focus();
	      return;
	}  
    if (!validarfecha(document.datos.valfecprev)){
	  	  document.datos.valfecprev.focus();
		  document.datos.valfecprev.select();
	      return;
    }  
    if (document.datos.monnro.value == ""){
	      alert("Debe selectar una moneda.");
		  document.datos.monnro.focus();
	      return;
    }  

    if (document.datos.pliqnro.value == ""){
	      alert("Debe selectar un período.");
	  	  document.datos.pliqnro.focus();
          return;
    }  
  
    if (document.datos.pliqdto.value == ""){
	      alert("Debe selectar un período de descuento.");
		  document.datos.pliqdto.focus();
          return;
	}  	

   //document.valida.location = "vales_liq_06.asp?pliqnro=" + document.datos.pliqdto.value + "&fecha1=" + document.datos.valfecped.value + "&fecha2=" + document.datos.valfecprev.value;		
   datosCorrectos();
}

function copiarAFecPrevista(){
  if (document.datos.valfecprev.value == ""){
      document.datos.valfecprev.value = document.datos.valfecped.value;
  }
}

function datosCorrectos(){
  abrirVentanaH('','vent_oculta',200,200); 
  document.datos.submit();
  document.all.empleg.focus();
  document.all.empleg.select();
}

function datosIncorrectos(){
  alert('Las fechas no se encuentran dentro del período de descuento.');
  document.datos.valfecped.focus();
  document.datos.valfecped.select();
}

function CambioTipoVale(){

<%'if l_tipo = "A" then%>
	if(document.datos.tvalenro2[document.datos.tvalenro2.selectedIndex].fijo == -1){
		document.datos.valmonto.value  = document.datos.tvalenro2[document.datos.tvalenro2.selectedIndex].monto;
		document.datos.valmonto.readOnly = true;
		//document.all.valmonto.className = 'deshabinp';
		document.all.valdesc.focus();
		document.all.valdesc.select();
	}
	else{
		document.datos.valmonto.value  = 0;
		document.datos.valmonto.readOnly = false;
		//document.all.valmonto.className = 'habinp';
		document.all.valmonto.focus();
		document.all.valmonto.select();		
	}	
<% 'end If %>
}

function Inicializar(){
	if(document.datos.tvalenro2[document.datos.tvalenro2.selectedIndex].fijo == -1){
		document.datos.valmonto.value  = document.datos.tvalenro2[document.datos.tvalenro2.selectedIndex].monto;
		document.datos.valmonto.readOnly = true;
		document.all.valdesc.focus();
		document.all.valdesc.select();
	}
	else{
		document.datos.valmonto.readOnly = false;
		document.all.valmonto.focus();
		document.all.valmonto.select();		
	}	
}

</script>

</head>
<body leftmargin="0" topmargin="0" rightmargin="0" bottommargin="0">

<form name="datos" action="vales_liq_03.asp?Tipo=<%=l_tipo%>"  target="vent_oculta" method="post">
<input type="hidden" name="tipo" value="<%=l_tipo%>">
<input type="hidden" name="valnroant" value="<%=l_valnro%>">
<input type="hidden" name="valnro" value="<%=l_valnro%>">
<input type="hidden" name="pliqnro" value="0">
<input type="Hidden" name="emplegant" value="<%= l_empleg %>">
<input type="hidden" name="ternro" value="<%=l_empleado%>">
<input type="hidden" name="valmonto2" value="<%=l_valmonto%>">
<input type="hidden" name="ppagnro" value="<%=l_ppagnro%>">
<input type="hidden" name="pronro" value="<%=l_pronro%>">
<input type="hidden" name="seleccion" value="">
<input type="hidden" name="seleccion1" value="">

<table border="0" cellpadding="0" cellspacing="0" width="100%" height="100%">
<tr style="border-color :CadetBlue;">
<td colspan="2" align="left" class="barra">Datos del Vale</td>
<td colspan="2" align="right" class="barra">
    <a tabindex="-1" class=sidebtnHLP href="Javascript:ayuda('<%= Request.ServerVariables("SCRIPT_NAME")%>');">Ayuda</a>
</td>	
</tr>
<tr>
    <td nowrap align="right"><b>Empleado:</b></td>
	<td colspan="3">
		<input id="empleg" type="text" maxlength="8" value="<%= l_empleg %>" size="8" name="empleg" onchange="javascript:buscarEmpleado();" onKeyPress="return Tecla(event.keyCode, 'emp')" class='deshabinp' readonly>
		<input tabindex="-1" class='deshabinp' readonly type="text" name="empleado" size="35" maxlength="35" value="<%= l_apnom%>">
	</td>
</tr>
<tr>
	<td align="right"><b>Monto:</b></td>
	<td colspan=3><input type="text" name="valmonto" size="20" maxlength="20" onkeypress="return Tecla(event.keyCode,'valor')" value="<%=l_valmonto%>">
	</td>
</tr>
<tr>
	<td align="right"><b>Descripci&oacute;n:</b></td>
	<td colspan=3><input type="text" name="valdesc" size="30" maxlength="30" value="<%=l_valdesc%>">
	</td>
</tr>
<tr>
	<td align="right">
      <b>Tipo Vale:</b>
    </td>
	<td align="left" colspan="3"> 	
	   <select name="tvalenro2" size="1" style="width:200px" onchange="CambioTipoVale();">
		<%	Set l_rs = Server.CreateObject("ADODB.RecordSet") 
		
  		    l_sql = "SELECT tvalenro, tvaledesabr, tvalemfijo, tvalemonto "
			l_sql  = l_sql  & "FROM tipovale "
			l_sql  = l_sql  & "ORDER BY tvaledesabr "			
			rsOpen l_rs, cn, l_sql, 0
			do until l_rs.eof		%>	
			<option fijo="<%= l_rs("tvalemfijo") %>" monto="<%= l_rs("tvalemonto") %>" value="<%= l_rs("tvalenro") %>" <%if CInt(l_rs("tvalenro")) = CInt(l_tvalenro2) then response.write "selected" end if%> > 
			<%= l_rs("tvaledesabr") %> </option>
		<%	    l_rs.Movenext
			loop
			l_rs.Close          %>	
		</select>
	</td>
</tr>
<tr>
    <td align="right"><b>Fec. Pedido:</b></td>
	<td>
		<input type="text" name="valfecped" maxlength="10" size="10" onchange="javascript:copiarAFecPrevista();" value="<%= l_valfecped%>">
		<a href="Javascript:Ayuda_Fecha(document.datos.valfecped);" tabindex="-1"><img src="/serviciolocal/shared/images/cal.gif" border="0"></a>	
	</td>
	<td align="right"><b>Fec. Prevista:</b></td>
	<td>
		<input type="text" name="valfecprev" maxlength="10" size="10" value="<%= l_valfecprev%>">
		<a href="Javascript:Ayuda_Fecha(document.datos.valfecprev);" tabindex="-1"><img src="/serviciolocal/shared/images/cal.gif" border="0"></a>	
	</td>
</tr>
<tr>
	<td align="right">
      <b>Moneda:</b>
    </td>
	<td align="left" colspan="3"> 	
	   <select name="monnro" size="1" style="width:200px">
		<%	Set l_rs = Server.CreateObject("ADODB.RecordSet") 
		
  		    l_sql = "SELECT monnro, mondesabr, monorigen "
			l_sql  = l_sql  & "FROM moneda "
			l_sql  = l_sql  & "ORDER BY mondesabr "
			rsOpen l_rs, cn, l_sql, 0
			do until l_rs.eof		%>	
			<option value="<%= l_rs("monnro") %>" <%if (CInt(l_rs("monnro")) = CInt(l_monnro)) OR ((CInt(l_monnro) = 0) AND (CInt(l_rs("monorigen")) = -1))then response.write "selected" end if%> > 
			<%= l_rs("mondesabr") %> </option>
		<%	    l_rs.Movenext
			loop
			l_rs.Close          %>	
		</select>
	</td>
</tr>
<tr>
	<td align="right">
      <b>Per&iacute;odo&nbsp;Dto.:</b>
    </td>
	<td align="left" colspan="3"> 	
	   <select name="pliqdto" size="1" style="width:200px">
		<%	Set l_rs = Server.CreateObject("ADODB.RecordSet") 
		
  		    l_sql = "SELECT pliqnro, pliqdesc "
			l_sql  = l_sql  & "FROM periodo "
			l_sql  = l_sql  & "ORDER BY pliqanio DESC, pliqmes DESC "
			rsOpen l_rs, cn, l_sql, 0
			do until l_rs.eof		%>	
			<option value="<%= l_rs("pliqnro") %>" <%if CInt(l_rs("pliqnro")) = CInt(l_pliqdto) then response.write "selected" end if%> > 
			<%= l_rs("pliqdesc") %> </option>
		<%	    l_rs.Movenext
			loop
			l_rs.Close          %>	
		</select>
	</td>
</tr>
<tr>
    <td align="right" class="th2" colspan=4>
	    <% call MostrarBoton ("sidebtnABM", "Javascript:Validar_Formulario();","Aceptar") %>	
		<a class=sidebtnABM href="Javascript:window.close()">Cancelar</a>
	</td>
</tr>
</table>
</form>

<iframe name="valida" src="blanc.asp" width="0" height="0"></iframe>
<script>Inicializar();</script>
</body>
</html>

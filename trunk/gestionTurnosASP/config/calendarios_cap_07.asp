<% Option Explicit %>
<!--#include virtual="/rhprox2/shared/db/conn_db.inc"-->
<!--#include virtual="/rhprox2/shared/inc/fecha.inc"-->
<!--#include virtual="/rhprox2/shared/inc/estado_evento.inc"-->
<!--
Archivo: calendarios_cap_02.asp
Descripción: Abm de Calendarios
Autor : Raul Chinestra
Fecha: 24/10/2003
-->
<% 
Dim l_evmonro
Dim l_evenro

'ADO
Dim l_tipo
Dim l_sql
Dim l_rs

l_tipo    = request.querystring("tipo")
l_tipo = "A"


l_evmonro = request.querystring("evmo")
l_evenro  = request.querystring("evenro")

%>
<html>
<head>
<link href="/rhprox2/shared/css/tables3.css" rel="StyleSheet" type="text/css">
<meta http-equiv="Content-Type" content="text/html; charset=iso-8859-1">
<title>Calendarios - Capacitación - RHPro &reg;</title>
</head>
<script src="/rhprox2/shared/js/fn_ayuda.js"></script>
<script src="/rhprox2/shared/js/fn_windows.js"></script>
<script src="/rhprox2/shared/js/fn_ay_generica.js"></script>
<script src="/rhprox2/shared/js/fn_fechas.js"></script>
<% If eventocerrado then %>
	<script src="/rhprox2/shared/js/fn_sololectura.js"></script>
	<script language="VBScript">
		function window_onload()
			sololectura()			
			document.datos.lugnro.disabled = true				
		end function
	</script>
<% End if %>


<script>

function Habilitar(opc){

switch (opc.value) {
   case '1' :
      document.datos.lu.disabled  = true;
	  document.datos.ma.disabled  = true;
	  document.datos.mi.disabled  = true;
	  document.datos.ju.disabled  = true;
	  document.datos.vi.disabled  = true;
	  document.datos.sa.disabled  = true;
	  document.datos.dom.disabled = true;	  	  
	  document.datos.semana.readOnly = true;
	  document.datos.semana.className="deshabinp" 
	  break;
   case '2' :
   	  document.datos.ma.disabled  = false;
      document.datos.lu.disabled  = false;
	  document.datos.mi.disabled  = false;
	  document.datos.ju.disabled  = false;
	  document.datos.vi.disabled  = false;	 
	  document.datos.sa.disabled  = false;
	  document.datos.dom.disabled = false;	  	  
	  document.datos.semana.readOnly = true;
	  document.datos.semana.className="deshabinp" 
	  break;
   case '3' :
      document.datos.lu.disabled  = false;
	  document.datos.ma.disabled  = false;
	  document.datos.mi.disabled  = false;
	  document.datos.ju.disabled  = false;
	  document.datos.vi.disabled  = false;
	  document.datos.sa.disabled  = false;
	  document.datos.dom.disabled = false;	  	  
	  document.datos.semana.readOnly = false;
	  document.datos.semana.className="habinp"
	  break;
} 

}

function Validar_Formulario()
{

if (document.datos.lugnro.value == 0) {
		alert('Debe Seleccionar un Lugar');document.datos.lugnro.focus();}
else if (!validarfecha(document.datos.calfecini)) {
    	document.datos.calfecini.focus();}
else if (!validarfecha(document.datos.calfecfin)) {
    	document.datos.calfecfin.focus();}
else if (!(menorque(document.datos.calfecini.value,document.datos.calfecfin.value))) {
		alert("La Fecha de Inicio debe ser menor o igual que la Fecha de Finalización.");document.datos.calfecini.focus();}		
else if (isNaN(document.datos.calhordes1.value)    ||
              (document.datos.calhordes1.value<0)  ||
			  (document.datos.calhordes1.value>23) ||
	     isNaN(document.datos.calhordes2.value)    ||
		      (document.datos.calhordes2.value<0)  ||
			  (document.datos.calhordes2.value>59) ||
			  (document.datos.calhordes1.value.length != 2) || 
			  (document.datos.calhordes2.value.length != 2)) {
			alert("Debe ingresar la hora desde o esta mal ingresada.");document.datos.calhordes1.focus();}		
else if (isNaN(document.datos.calhorhas1.value)    ||
              (document.datos.calhorhas1.value<0)  || 
			  (document.datos.calhorhas1.value>23) ||
	     isNaN(document.datos.calhorhas2.value)    ||
		      (document.datos.calhorhas2.value<0)  ||
			  (document.datos.calhorhas2.value>59) ||
			  (document.datos.calhorhas1.value.length != 2) || 
			  (document.datos.calhorhas2.value.length != 2)) {
			alert("Debe ingresar la hora hasta o esta mal ingresada.");	document.datos.calhorhas1.focus();}						
else if ( document.datos.calhordes1.value>document.datos.calhorhas1.value ||
		 (document.datos.calhordes1.value==document.datos.calhorhas1.value &&
		  document.datos.calhordes2.value>document.datos.calhorhas2.value)) {
			alert("La Hora Desde debe ser Menor que la Hora Hasta."); 	document.datos.calhordes1.focus();}													
else if (
        (document.datos.rbopc(1).checked || document.datos.rbopc(2).checked) && 
		 document.datos.lu.checked  == false &&
		 document.datos.ma.checked  == false &&
		 document.datos.mi.checked  == false &&
		 document.datos.ju.checked  == false &&
		 document.datos.sa.checked  == false &&
		 document.datos.dom.checked == false &&
		 document.datos.vi.checked  == false ) {
	    alert('Debe Seleccionar al menos un día de la Semana');document.datos.lu.focus(); }			
else if (document.datos.rbopc(2).checked && 
		 document.datos.semana.value <= 0 ||  document.datos.semana.value >= 5 ) {
	    alert('Debe ingresar la Semana o esta mal ingresada');document.datos.semana.focus(); }					
else {
		//var d=document.datos;
		//document.valida.location = "calendarios_cap_09.asp?tipo=<%'= l_tipo%>&temnro="+document.datos.calnro.value + "&temdesabr="+document.datos.temdesabr.value;	
		valido();
        //parent.ifrm.location.reload();
		//parent.ifrm2.location.reload();
		
}

}

function valido(){  
  abrirVentanaH('','calen_mas',300,300);
  document.datos.target='calen_mas';
  document.datos.action="calendarios_cap_08.asp?tipo=<%= l_tipo %>&evmo=<%= l_evmonro%>";
  document.datos.submit();
}

function invalido(texto){
  alert(texto);
}

function Nuevo_Dialogo(w_in, pagina, ancho, alto)
{
 return w_in.showModalDialog(pagina,'', 'center:yes;dialogWidth:' + ancho.toString() + ';dialogHeight:' + alto.toString() + ';');
}

function Ayuda_Fecha(txt){
 var jsFecha = Nuevo_Dialogo(window, '/rhprox2/shared/js/calendar.html', 16, 15);
 if (jsFecha == null){
 	//txt.value = '';
 }else{
 	txt.value = jsFecha;
 	//DiadeSemana(jsFecha);
	}
}

</script>

<body leftmargin="0" rightmargin="0" topmargin="0" bottommargin="0"  onload="javascript:document.datos.lugnro.focus()">
<form name="datos" action="calendarios_cap_08.asp?tipo=<%= l_tipo %>&evmo=<%= l_evmonro%>" method="post" >
<input type="Hidden" name="calnro" value="<%= l_calnro %>">

<table cellspacing="0" cellpadding="0" border="0" width="100%" height="100%" style="border: thin solid Gray;">
<tr>
	<td class="barra">
	 <a class=sidebtnSHW href="Javascript:abrirVentana('rep_cronograma_eventos_cap_00.asp','',800,550)">Reporte Eventos</a>
	 </td>
    <td colspan="4"  align="center" class="barra">Asignaci&oacute;n Masiva de Calendarios</td>
</tr>
<tr>
	<td colspan="5" height="1px"></td>
</tr>

<tr>
    <td align="right"><b>Lugar:</b></td>
	<td colspan="4">
		<select style="width:385px" name=lugnro size="1">
		<% if l_tipo = "A" then %> 
			<option value=0>«Seleccione una Opción»</option>
		<% end if %> 
		
		<%	
			Set l_rs = Server.CreateObject("ADODB.RecordSet")
			l_sql = "SELECT lugnro, lugdesabr"
			l_sql  = l_sql  & " FROM cap_lugar "
			l_sql  = l_sql  & " ORDER BY lugdesabr"
			rsOpen l_rs, cn, l_sql, 0
			do until l_rs.eof		%>	
			<option value= <%= l_rs("lugnro") %> > 
			<%= l_rs("lugdesabr") %> (<%=l_rs("lugnro")%>) </option>
		<%			l_rs.Movenext
			loop
			l_rs.Close %>	
		</select>
		<script> document.datos.lugnro.value=0</script>
	</td>	
</tr>

<tr>
	<td align="right"><b>Fecha Inicio:</b></td>
	<td>
		<input  type="text" name="calfecini" size="10" maxlength="10" value="" >
		<a href="Javascript:Ayuda_Fecha(document.datos.calfecini);"><img src="/rhprox2/shared/images/cal.gif" border="0"></a>
	</td>
	<td align="right"><b>Fecha Finalización:</b></td>
	<td>
		<input  type="text" name="calfecfin" size="10" maxlength="10" value="">
		<a href="Javascript:Ayuda_Fecha(document.datos.calfecfin);"><img src="/rhprox2/shared/images/cal.gif" border="0"></a>
	</td>
</tr>

<tr>
	<td align="right"><b>Hora Desde:</b></td>
	<td>
	<input type="text" name="calhordes1" size="2" maxlength="2" >
	<b>:</b>
    <input type="text" name="calhordes2" size="2" maxlength="2" >
	</td>
	<td align="right"><b>Hora Hasta:</b></td>
	<td>
	<input type="text" name="calhorhas1" size="2" maxlength="2" >
	<b>:</b>
    <input type="text" name="calhorhas2" size="2" maxlength="2" >
	</td>
</tr>


<tr>
    <td align="left">
	<table border="0" cellpadding="0" cellspacing="0"> 
	  <tr> 
	  	<td> 	<input type=radio name=rbopc value=1 CHECKED onclick="Habilitar(this)"> <b>Dias Corridos</b><br>
		</td>
	  </tr>
	  <tr> 
	  	<td> 	<input type=radio name=rbopc value=2 onclick="Habilitar(this)"> <b>Dias Semana</b><br>
		</td>
	  </tr>
	  <tr> 
	  	<td> 	<input type=radio name=rbopc value=3 onclick="Habilitar(this)"> <b>Dias Mes</b><br>
		</td>
	  </tr>
	</table>
	</td>
	
	<td align="left">
	<table> 
	  <tr>
	      <td>LU</td>
	 	  <td>MA</td>
	  	  <td>MI</td>
	  	  <td>JU</td>
	  	  <td>VI</td>
  	  	  <td>SA</td>
	  	  <td>DO</td>
	  </tr>	
	  <tr>
	      <td><input disabled type=checkbox name=lu> </td>
	 	  <td><input disabled type=checkbox name=ma> </td>
	  	  <td><input disabled type=checkbox name=mi> </td>
	  	  <td><input disabled type=checkbox name=ju> </td>
	  	  <td><input disabled type=checkbox name=vi> </td>
	  	  <td><input disabled type=checkbox name=sa> </td>
	  	  <td><input disabled type=checkbox name=dom> </td>
		  
	  </tr>	
 	</table> 
	</td>
	
	<td> 
	<table> 
	  <tr>
	      <td align="center">Semana</td>
      </tr>	  
  	  <tr>	
        <td align="center">
	    <input readOnly class="deshabinp" type="text" name="semana" size="1" maxlength="1" >
		</td>
      </tr>
    </table>	  	
		
	</td> 

	<td> &nbsp;
		<a class=sidebtnABM href="Javascript:Validar_Formulario()">Crear Calendario</a>
		<a class=sidebtnABM href="Javascript:window.location.reload()">Limpiar</a>
		<iframe name="valida2" frameborder="0" style="display:none" src="" width="0" height="0"></iframe> 
	</td> 
</tr>

</table>

</form>
<%
set l_rs = nothing
'l_Cn.Close
'set l_Cn = nothing
%>
</body>
</html>

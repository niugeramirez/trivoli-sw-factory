<%Option Explicit %>
<!--#include virtual="/serviciolocal/shared/inc/sec.inc"-->
<!--#include virtual="/serviciolocal/shared/inc/const.inc"-->
<!--#include virtual="/serviciolocal/shared/db/conn_db.inc"-->
<% 
' Variables
Dim l_btc_tmp_esp_no_resp
Dim l_btc_tmp_esp_sin_prog
Dim l_btc_tmp_lect_reg
Dim l_btc_tmp_dorm
Dim l_btc_usa_reg
Dim l_btc_max_proc
Dim l_btc_mult_logs
Dim l_btc_path_proc
Dim l_btc_path_logs
Dim l_btc_form_fecha

dim l_tipo

dim l_rs
dim l_rs1
dim l_sql

l_tipo      = Request.QueryString("tipo")
l_caudnro   = Request.QueryString("caudnro")

Set l_rs = Server.CreateObject("ADODB.RecordSet")
    l_sql = "SELECT btc_tmp_esp_no_resp, btc_tmp_esp_sin_prog, btc_tmp_lect_reg, btc_tmp_dorm, btc_usa_reg, "
    l_sql = l_sql & " btc_max_proc, btc_mult_logs, btc_path_proc, btc_path_logs, btc_form_fecha"  
	l_sql = l_sql & " FROM  batch_config"

		l_rs.MaxRecords = 1
		rsOpen l_rs, cn, l_sql, 0 
		if not l_rs.eof then
		   l_btc_tmp_esp_no_resp  = l_rs("btc_tmp_esp_no_resp")
		   l_btc_tmp_esp_sin_prog = l_rs("btc_tmp_esp_sin_prog")
		   l_btc_tmp_lect_reg     = l_rs("btc_tmp_lect_reg")
		   l_btc_tmp_dorm         = l_rs("btc_tmp_dorm")
		   l_btc_usa_reg          = l_rs("btc_usa_reg")
		   l_btc_max_proc         = l_rs("btc_max_proc")
		   l_btc_mult_logs        = l_rs("btc_mult_logs")
		   l_btc_path_proc        = l_rs("btc_path_proc")
		   l_btc_path_logs        = l_rs("btc_path_logs")
		   l_btc_form_fecha       = l_rs("btc_form_fecha")
		else
		   l_btc_tmp_esp_no_resp  = 5
		   l_btc_tmp_esp_sin_prog = 5
		   l_btc_tmp_lect_reg     = 1
		   l_btc_tmp_dorm         = 10
		   l_btc_usa_reg          = 0
		   l_btc_max_proc         = 2
		   l_btc_mult_logs        = -1
		   l_btc_path_proc        = "c:\visual\exe"
		   l_btc_path_logs        = "c:\logs"
		   l_btc_form_fecha       = "MM/DD/YYYY"
		end if
		l_rs.Close
		set l_rs = nothing
%>

<html>
<head>
<link href="/serviciolocal/shared/css/tables3.css" rel="StyleSheet" type="text/css">
<meta http-equiv="Content-Type" content="text/html; charset=iso-8859-1">
<title><%= Session("Titulo")%>Configuraci&oacute;n de Batch Proceso</title>
<script src="/serviciolocal/shared/js/fn_windows.js"></script>
<script src="/serviciolocal/shared/js/fn_confirm.js"></script>
<script src="/serviciolocal/shared/js/fn_hora.js"></script>
<script src="/serviciolocal/shared/js/fn_ayuda.js"></script>
<script src="/serviciolocal/shared/js/fn_numeros.js"></script>
<script src="/serviciolocal/shared/js/fn_fechas.js"></script>
<script src="/serviciolocal/shared/js/fn_ayuda.js"></script>
<script src="/serviciolocal/shared/js/fn_fechas.js"></script>
<script>
function Validar_Formulario()
{
if (document.datos.tenr.value == "" ){
	alert("Ingrese el Tiempo de Espera No Responde.");
	document.datos.tenr.focus();}
else	
if (!(validanumero(document.datos.tenr,15,0))) {
		alert("El Tiempo de Espera No Responde debe ser entero.");
		document.datos.tenr.focus();
	}
else	
if (document.datos.tesp.value == "" ){
	alert("Ingrese el Tiempo de Espera Sin Progreso.");
	document.datos.tesp.focus();}
else
if (!(validanumero(document.datos.tesp,15,0))) {
		alert("El Tiempo de Espera Sin Progreso debe ser entero.");
		document.datos.tesp.focus();
	}
else
if (document.datos.tldr.value == "" ){
	alert("Ingrese el Tiempo de Lectura de Registraciones.");
	document.datos.tldr.focus();}
else
if (!(validanumero(document.datos.tldr,15,0))) {
		alert("El Tiempo de Lectura de Registraciones debe ser entero.");
		document.datos.tldr.focus();
	}
else
if (document.datos.tdd.value == "" ){
	alert("Ingrese el Tiempo de Dormida.");
	document.datos.tdd.focus();}
else
if (!(validanumero(document.datos.tdd,15,0)) || parseInt(document.datos.tdd.value) <= 5) {
		alert("El Tiempo de Dormida debe un ser entero mayor a 5.");
		document.datos.tdd.focus();
	}
else
if (document.datos.mproc.value == "" ){
	alert("Ingrese el Máximo Nro. de Procesos Concurrentes.");
	document.datos.mproc.focus();}
else
if (!(validanumero(document.datos.mproc,15,0))) {
		alert("El Máximo Nro. de Procesos Concurrentes debe ser entero.");
		document.datos.mproc.focus();
	}
else
if (document.datos.pproc.value == "" ){
	alert("Ingrese el Path para el Proceso.");
	document.datos.pproc.focus();}
else
if (document.datos.pproc.value.length > 100){
	alert("EL Path para el Proceso no puede superar los 100 caracteres.");		
	document.datos.pproc.focus();}
else	
if (document.datos.plogs.value == "" ){
	alert("Ingrese el Path para el LOG.");
	document.datos.plogs.focus();}
else
if (document.datos.plogs.value.length > 100){
	alert("EL Path para los LOGS no puede superar los 100 caracteres.");		
	document.datos.plogs.focus();}
else	
if (document.datos.fecha.value == "" ){
	alert("Ingrese el formato de la Fecha.");
	document.datos.fecha.focus();}
else
  { document.datos.submit();
  }

}

function Ayuda_Fecha(txt){
 var jsFecha = Nuevo_Dialogo(window, '/serviciolocal/shared/js/calendar.html', 16, 15);

 if (jsFecha == null) txt.value = ''
 else txt.value = jsFecha;
}

function cambioRegistracion(){
var s;

    if (document.datos.ureg.checked)
	   s = "habinp";	
	else
	   s = "deshabinp";   
   
   document.datos.tldr.className=s; 
  if (document.datos.ureg.checked){
     document.datos.ureg.value = "-1";
	 document.datos.tldr.disabled = false;
	 }
	 else 
	 {  document.datos.ureg.value = "0";
	    document.datos.tldr.disabled = true;
		document.datos.tldr.value = 0;
	 }
}

function cambioArchivo(){

  if (document.datos.march.checked)
     document.datos.march.value = "-1"
	 else document.datos.march.value = "0"
}

</script>

</head>
<body leftmargin="0" topmargin="0" rightmargin="0" bottommargin="0" onload="cambioRegistracion();cambioArchivo();">

<table border="0" cellpadding="0" cellspacing="0">
<tr style="border-color :CadetBlue;">
<td colspan="4" class="th2" align="right">		  
<a class=sidebtnHLP href="Javascript:ayuda('<%= Request.ServerVariables("SCRIPT_NAME")%>');">Ayuda</a>
</td>
</tr>

<form name="datos" action="batch_config_01.asp" method="post">
<input type="hidden" name="tipo" value="<%=l_tipo%>">
<tr>
	<td align="right"><b>Tiempo de Espera No Responde:</b></td>
	<td colspan=3><input type="text" name="tenr" size="10" maxlength="10" value="<%=l_btc_tmp_esp_no_resp%>" >
	<b>(Minutos)</b>
	</td>
</tr>
<tr>
	<td align="right"><b>Tiempo de Espera Sin Progreso:</b></td>
	<td colspan=3><input type="text" name="tesp" size="10" maxlength="10" value="<%=l_btc_tmp_esp_sin_prog%>">
	<b>(Minutos)</b>
	</td>
</tr>
<tr>
	<td align="right"><b>Tiempo de Dormida:</b></td>
	<td colspan=3><input type="text" name="tdd" size="10" maxlength="10" value="<%=l_btc_tmp_dorm%>">
	<b>(Segundos)</b></td>
</tr>
<tr>
	<td align="right">
		<input type="checkbox" name="ureg" onclick="cambioRegistracion(this);" value="<%= l_btc_usa_reg %>" <% if l_btc_usa_reg = "-1" then %> checked <% end if %>>
	</td>
	<td align="left" colspan=3><b>Usa Lectura de Registraciones</b></td>
</tr>
<tr>
	<td align="right"><b>Tiempo de Lectura de Registraciones:</b></td>
	<td colspan=3><input type="text" name="tldr" size="10" maxlength="10" value="<%=l_btc_tmp_lect_reg%>">
	<b>(Minutos)</b>
	</td>
</tr>
<tr>
	<td align="right"><b>Máximo Nro de Procesos Concurrentes:</b></td>
	<td colspan=3><input type="text" name="mproc" size="10" maxlength="10" value="<%=l_btc_max_proc%>">
	</td>
</tr>
<tr>
	<td align="right">
		<input type="checkbox" name="march" onclick="cambioArchivo();" value="<%= l_btc_mult_logs %>" <% if l_btc_mult_logs = "-1" then %> checked <% end if %>>
	</td>
	<td align="left" colspan=3><b>Genera Múltiples Archivos de LOG</b></td>
</tr>
<tr>
	<td align="right"><b>Procesos:</b></td>
	<td colspan=3><input type="text" name="pproc" size="30" value="<%=l_btc_path_proc%>">
	</td>
</tr>
<tr>
	<td align="right"><b>Flog:</b></td>
	<td colspan=3><input type="text" name="plogs" size="30" value="<%=l_btc_path_logs%>">
	</td>
</tr>
<tr>
	<td align="right"><b>Formato Fecha:</b></td>
	<td colspan=3>
		<select name="fecha" size="1">
			<option value="MM/DD/YYYY">MM/DD/YYYY</option>
			<option value="MM/DD/YY"  >MM/DD/YY</option>
			<option value="DD/MM/YYYY">DD/MM/YYYY</option>				
			<option value="DD/MM/YY"  >DD/MM/YY</option>		
		</select> 
	<script>document.datos.fecha.value = "<%=l_btc_form_fecha%>"</script>	
	</td>
</tr>

</form>
</table>

<table width="100%" border="0" cellspacing="0" cellpadding="0">
<tr>
    <td align="right" class="th2">
		<a class=sidebtnABM href="Javascript:Validar_Formulario()">Aceptar</a>
		<a class=sidebtnABM href="Javascript:window.close()">Cancelar</a>
	</td>
</tr>
</table>

</body>
</html>

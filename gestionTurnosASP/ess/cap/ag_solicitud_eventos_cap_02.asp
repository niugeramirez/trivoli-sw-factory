<% Option Explicit %>
<!--#include virtual="/serviciolocal/ess/shared/inc/const.inc"-->
<!--#include virtual="/serviciolocal/shared/db/conn_db.inc"-->
<!--#include virtual="/serviciolocal/shared/inc/fecha.inc"-->
<!--#include virtual="/serviciolocal/ess/shared/inc/accesoESS.inc"-->
<!--
Archivo: ag_solicitud_eventos_cap_00.asp
Descripción: Abm de Solicitud de eventos
Autor : Raul Chinestra
Fecha: 30/03/2004
-->
<% 
on error goto 0

'Datos del formulario
Dim l_solnro
Dim l_soldesabr
Dim l_soldesext
Dim l_soldurdias

'ADO
Dim l_tipo
Dim l_sql
Dim l_rs

Dim l_ternro
Dim l_solfec

l_tipo = request.querystring("tipo")
l_ternro = l_ess_ternro

%>
<html>
<head>
<link href="../<%= c_estilo %>" rel="StyleSheet" type="text/css">
<meta http-equiv="Content-Type" content="text/html; charset=iso-8859-1">
<title>Solicitud de Eventos - Capacitación - RHPro &reg;</title>
</head>
<script src="/serviciolocal/shared/js/fn_ayuda.js"></script>
<script src="/serviciolocal/shared/js/fn_windows.js"></script>
<script src="/serviciolocal/shared/js/fn_fechas.js"></script>
<script>
function Validar_Formulario()
{
if (document.datos.soldesabr.value == "") {
	alert("Debe ingresar la Descripción Abreviada.");
	document.datos.soldesabr.focus();
}	
else	
    {
	if (validarfecha(document.datos.solfec)) {
		if ( (isNaN(document.datos.soldurdias.value))) {
	    	alert("La duración pretendida es Incorrecta");	
			document.datos.soldurdias.focus();
		}	
		else
			{
			 if (document.datos.soldesext.value.length > 200) {
		  	   alert("La Descripción Extendida es demasiado grande.");		
	           document.datos.soldesext.focus();
			 }
			 else {
			    	var d=document.datos;
				    document.valida.location = "ag_solicitud_eventos_cap_06.asp?tipo=<%= l_tipo%>&solnro="+document.datos.solnro.value + "&soldesabr="+document.datos.soldesabr.value;	
			}		
		}
	}	
}
}

function valido(){
  document.datos.submit();
}

function invalido(texto){
  alert(texto);
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

</script>
<% 
Set l_rs = Server.CreateObject("ADODB.RecordSet")
select Case l_tipo
	Case "A":
		l_soldesabr  = ""
		l_solfec  = ""
		l_solnro     = 0
		l_soldesext = ""
		l_soldurdias = ""
	Case "M":
		l_solnro = request.querystring("cabnro")
		l_sql = "SELECT solnro, soldesabr, soldesext, soldurdias, solfec "
		l_sql  = l_sql  & " FROM cap_solicitud "
		l_sql  = l_sql  & " WHERE solnro = " & l_solnro
		rsOpen l_rs, cn, l_sql, 0 
		if not l_rs.eof then
			l_soldesabr  = l_rs("soldesabr")
			l_solfec 	 = l_rs("solfec")
			l_soldesext  = l_rs("soldesext")
			l_solnro     = l_rs("solnro")
			l_soldurdias = l_rs("soldurdias")			
		end if
		l_rs.Close
end select
%>
<body leftmargin="0" rightmargin="0" topmargin="0" bottommargin="0" onload="javascript:document.datos.soldesabr.focus()">
<form name="datos" action="ag_solicitud_eventos_cap_03.asp?tipo=<%= l_tipo %>&ternro=<%= l_ternro %>" method="post" >
<input type="Hidden" name="solnro" value="<%= l_solnro %>">

<table cellspacing="0" cellpadding="0" border="0" width="100%"  height="100%">
<tr>
    <td class="th2">Datos de la Solicitud</td>
	<td colspan="3" class="th2" align="right">		  
		&nbsp;
	</td>
</tr>

<tr>
</tr>

<tr>
    <td align="right"><b>Desc. Abreviada:</b></td>
	<td colspan="3">
		<input type="text" name="soldesabr" size="60" maxlength="50" value="<%= l_soldesabr %>">
	</td>
</tr>
<tr>
	<td align="right"><b>Fecha Aprox.:</b></td><td>
	<input  type="text" name="solfec" size="10" maxlength="10" value="<%= l_solfec %>"   >
	<a href="Javascript:Ayuda_Fecha(document.datos.solfec);"><img src="/serviciolocal/shared/images/cal.gif" border="0"></a>
	</td>
</tr>

<tr>
    <td align="right"><b>Duración:</b></td>
	<td>
		<input type="text" name="soldurdias" size="3" maxlength="3" value="<%= l_soldurdias %>">&nbsp; <b>Días</b>
	</td>	
</tr>
<tr>
    <td align="right"><b>Desc. Extendida:</b></td>
	<td colspan="3" align="left">
	    <textarea name="soldesext" rows="3" cols="45" maxlength="200"><%=trim(l_soldesext)%></textarea>
	</td>
</tr>

<tr>
    <td  colspan="4" align="right" class="th2">
		<a class=sidebtnABM href="Javascript:Validar_Formulario()">Aceptar</a>
		<a class=sidebtnABM href="Javascript:window.close()">Cancelar</a>
	</td>
</tr>
</table>

<iframe name="valida" style="visibility=hidden;"  src="blanc.asp" width="100%" height="50%"></iframe> 
</form>
<%
set l_rs = nothing
cn.Close
set Cn = nothing
%>
</body>
</html>

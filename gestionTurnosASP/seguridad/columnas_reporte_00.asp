<%Option Explicit %>
<!--#include virtual="/serviciolocal/shared/inc/sec.inc"-->
<!--#include virtual="/serviciolocal/shared/inc/const.inc"-->
<!--#include virtual="/serviciolocal/shared/db/conn_db.inc"-->
<!--
-----------------------------------------------------------------------------
Archivo: columnas_reporte_00.asp
Descripcion: Modulo que se encarga de los abm de las columnas de un
             reporte realizado con ConfRep.
Modificacion:
    29/07/2003 - Scarpa D. - Agregado de la columna confsuma   
    25/08/2003 - Scarpa D. - Agregado de la columna confval2
    14/04/2004 - Alvaro Bayon - maxrecords!!!. Botón de ayuda
-----------------------------------------------------------------------------
-->

<% 
' Variables

Dim l_sql
Dim l_rs
Dim l_repnro

l_repnro = Request.QueryString("repnro")

if len(l_repnro) = 0 or l_repnro = 0 then
	Response.write "<script>alert('Seleccione un reporte.');window.close();</script>"
end if

%>

<html>
<head>
<link href="/serviciolocal/shared/css/tables3.css" rel="StyleSheet" type="text/css">
<meta http-equiv="Content-Type" content="text/html; charset=iso-8859-1">
<title><%= Session("Titulo")%>Configuraci&oacute;n de Reportes</title>
<script src="/serviciolocal/shared/js/fn_windows.js"></script>
<script src="/serviciolocal/shared/js/fn_confirm.js"></script>
<script src="/serviciolocal/shared/js/fn_ayuda.js"></script>
<script>
function Recargar()
{
	document.ifrm.location.href= 'columnas_reporte_01.asp?repnro=' + document.datos.repnro.value ;
}

function Alta(){
  var param = '?Tipo=A&repnro='+document.datos.repnro.value;

  abrirVentana('columnas_reporte_02.asp' + param,'',545,260);
}

function Modificar(){
  var param = '?Tipo=M&repnro='+document.datos.repnro.value + '&confnrocol=' + document.ifrm.datos.cabnro.value+ '&conftipo=' + document.ifrm.datos.conftipo.value+ '&confetiq=' + document.ifrm.datos.confetiq.value+ '&confval=' + document.ifrm.datos.confval.value + '&confval2=' + document.ifrm.datos.confval2.value+ '&confaccion=' + document.ifrm.datos.confaccion.value;  
  abrirVentanaVerif('columnas_reporte_02.asp' + param,'',545,260);
}
    	   
</script>

</head>
<body leftmargin="0" topmargin="0" rightmargin="0" bottommargin="0">
<table border="0" cellpadding="0" cellspacing="0" height="100%">
<tr style="border-color :CadetBlue;">
<td colspan="1" align="left" class="barra">Configuraci&oacute;n de Reportes</td>
<td colspan="1" align="right" class="barra">
	<a class=sidebtnABM href="Javascript:Alta();">Alta</a>
	<a class=sidebtnABM href="Javascript:eliminarRegistro(document.ifrm,'columnas_reporte_04.asp?repnro='+document.datos.repnro.value + '&confnrocol=' + document.ifrm.datos.cabnro.value+ '&conftipo=' + document.ifrm.datos.conftipo.value+ '&confetiq=' + document.ifrm.datos.confetiq.value+ '&confval=' + document.ifrm.datos.confval.value + '&confval2=' + document.ifrm.datos.confval2.value + '&confaccion=' + document.ifrm.datos.confaccion.value)">Baja</a>
	<a class=sidebtnABM href="Javascript:Modificar();">Modifica</a>	
</td>
<td class="th2" align="right">
	  <a class=sidebtnHLP href="Javascript:ayuda('<%= Request.ServerVariables("SCRIPT_NAME")%>');">Ayuda</a>
</td>
</tr>

<form name=datos>
<%
' BUSCAR PERIODOS PARA EL <SELECT>
  Set l_rs = Server.CreateObject("ADODB.RecordSet")
  l_sql = "SELECT reporte.repnro, reporte.repdesc "
  l_sql = l_sql & "FROM reporte "
  
  rsOpen l_rs, cn, l_sql, 0 
%>
<tr>
    <td align="right"><b>Reporte:</b></td>
	<td colspan="2">
		<select name="repnro" onChange="Recargar();">
		<%l_rs.MoveFirst
		 do while not l_rs.eof%>
			<option value=<%=l_rs("repnro")%>><%=l_rs("repdesc")%>(<%=l_rs("repnro")%>)</option>
		<%l_rs.MoveNext
		loop
		l_rs.Close
		set l_rs = nothing%>
		</select>
		<script>document.datos.repnro.value='<%=l_repnro%>'</script>
	</td>
</tr>

</form>

<tr valign="top">
   <td colspan="3" style="" height="100%">
   <iframe name="ifrm" src="columnas_reporte_01.asp?repnro=<%=l_repnro%>" width="100%" height="100%"></iframe> 
   </td>
</tr>

</table>
</body>
</html>

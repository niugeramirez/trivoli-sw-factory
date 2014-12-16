<% Option Explicit %>
<!--#include virtual="/serviciolocal/ess/shared/inc/const.inc"-->
<!--#include virtual="/serviciolocal/shared/inc/sec.inc"-->
<!--#include virtual="/serviciolocal/shared/inc/const.inc"-->
<!-- 
'================================================================================
'Archivo		: recibosueldo_ess_00.asp
'Descripción	: Muestra los Recibos de Sueldos del Empleado
'Autor			: GdeCos
'Fecha			: 14-04-2005
'Modificado		: 11/08/2006 - Martin Ferraro - Paso como parametro el asp del recibo de sueldo a mostrar
'================================================================================
 -->
<%
on error goto 0

Dim l_salida
Dim l_aux
Dim l_car
Dim l_legajo_param

l_aux = request.querystring("salida")

'Si tiene por parametro un asp para la salida del recibo debo acomodar parametros porque tiene
'formato del estilo xxxxxxx?empleg= donde xxxx es el nombre del asp de salida
if trim(l_aux) <> "" then
	l_car = inStr(l_aux,"?")
	if l_car > 0 then
		l_salida = mid(l_aux,1,l_car-1)
		l_legajo_param = mid(l_aux,l_car+8,len(l_aux))
		
	else
		response.write "Error en parametros."
		response.end
	end if
else
	l_salida = ""
	l_legajo_param = request.querystring("empleg")
end if
%>
<html>
<head>
<link href="../<%= c_estilo %>" rel="StyleSheet" type="text/css">
<title>Recibos de Sueldos - Autogestion - RHPro &reg;</title>
<script src="/serviciolocal/shared/js/fn_ayuda.js"></script>
<script>
function Recibo()
{
	if (document.ifrm.datos.cabnro.value != 0)
	{
		//window.location = "rep_recibo_ess_00.asp?rec="+document.ifrm.datos.cabnro.value+"&empleg=<%'= l_empleg%>";
		window.location = "rep_recibo_ess_00.asp?rec=" + document.ifrm.datos.cabnro.value + "&empleg=<%= l_legajo_param%>&salida=<%= l_salida %>";

	}
	else
	{
		alert('Debe seleccionar un recibo');
	}
}
</script>
</head>

<body>
      <table border="0" cellpadding="0" cellspacing="0" height="100%">
        <tr>
          <th align="left">Recibos de Sueldos</th>
          <th nowrap align="right">
	  	  	 <a href="javascript:Recibo();" class="sidebtnSHW">Ver Recibo</a>			  
		  </th>
        </tr>
        <tr valign="top" height="100%">
          <td colspan="2" style="">
      	  <iframe name="ifrm" src="recibosueldo_ess_01.asp?empleg=<%= l_legajo_param%>" width="100%" height="100%"></iframe> 
	      </td>
        </tr>
      </table>
</body>
</html>

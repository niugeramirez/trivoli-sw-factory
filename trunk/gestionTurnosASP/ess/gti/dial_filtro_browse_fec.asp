<% Option Explicit %>
<!--#include virtual="/serviciolocal/ess/shared/inc/const.inc"-->
<%
Dim l_campo 
Dim l_fuente 
Dim l_orden 

'esto es para pasar el parametro a la funcion que formatea la fecha
Dim l_base
l_base = Session("base")

l_campo  = request.querystring("campo")
l_fuente = request.querystring("fuente")
l_orden  = request.querystring("orden")
%>
<html>
<head>
<link href="../<%= c_estilo %>" rel="StyleSheet" type="text/css">
<meta http-equiv="Content-Type" content="text/html; charset=iso-8859-1">
<title>Filtrar - Gesti&oacute;n de Tiempos - RHPro &reg;</title>
</head>
<script src="/serviciolocal/shared/js/fn_fechas.js"></script>
<script>
window.returnValue='';

function filtrar()
{
  for (i=0;i<=2;i++)
  {
    if (document.datos.orden[i].checked)
	  var sel = document.datos.orden[i].value
  }
  if (document.datos.texto.value == "")
    alert("Debe ingresar un valor")
  else
    if (validarfecha(document.datos.texto)) 
    {
	  if (sel == "1")	
	    {
  	    var txt = '<%= l_campo %> < ' + cambiafecha(document.datos.texto.value,true,<%=l_base%>) + ' '
		}
	  if (sel == "2")	
	    {
  	    var txt = '<%= l_campo %> > ' + cambiafecha(document.datos.texto.value,true,<%=l_base%>) + ' '
		}
	  if (sel == "3")	
	    {
  	    var txt = '<%= l_campo %> = ' + cambiafecha(document.datos.texto.value,true,<%=l_base%>) + ' '
		}
      window.returnValue = txt; 
	  window.close();  
    }
}
window.resizeTo(300,180)
</script>

<body leftmargin="0" rightmargin="0" topmargin="0" bottommargin="0">
<form name="datos" method="post">
<table cellspacing="1" cellpadding="0" border="0" width="100%">
  <tr>
    <td class="th2" colspan="3">Filtrar</td>
  </tr>
<tr>
    <td align="right"><b>Anterior a:</b></td>
	<td><input type="Radio" name="orden" value="1" checked></td>
</tr>
<tr>
    <td align="right"><b>Posterior a:</b></td>
	<td><input type="Radio" name="orden" value="2"></td>
</tr>
<tr>
    <td align="right"><b>Entre el:</b></td>
	<td><input type="Radio" name="orden" value="3"></td>
</tr>
<tr>
    <td align="right"><b>y el:</b></td>
	<td><input type="Text" name="texto" value=""></td>
</tr>
</table>
<table width="100%" border="0" cellspacing="0" cellpadding="0">
<tr>
    <td align="right" class="th2">
		<a class=sidebtnABM href="#" onclick="Javascript:filtrar()">Aceptar</a>
		<a class=sidebtnABM href="#" onclick="Javascript:window.close()">Cancelar</a>		
	</td>
</tr>
</table>
</form>
</body>
</html>

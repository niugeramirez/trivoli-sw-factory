<% Option Explicit %>
<%
Dim l_campo 
Dim l_fuente 
Dim l_orden 

l_campo  = request.querystring("campo")
l_fuente = request.querystring("fuente")
l_orden  = request.querystring("orden")
%>
<html>
<head>
<link href="/turnos/shared/css/tables3.css" rel="StyleSheet" type="text/css">
<meta http-equiv="Content-Type" content="text/html; charset=iso-8859-1">
<title><%= Session("Titulo")%>Filtrar - RHPro &reg;</title>
</head>
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
  
  if (parseFloat(document.datos.texto.value) > 99999999999999999999)
 	alert("El n�mero es demasiado grande.");
  else  
    if (isNaN(document.datos.texto.value)) 
	  alert("El valor debe ser num�rico.");
    else  
    {
	  if (sel == "1")	
	    {
  	    var txt = '<%= l_campo %> > ' + document.datos.texto.value + ' '
		}
	  if (sel == "2")	
	    {
  	    var txt = '<%= l_campo %> < ' + document.datos.texto.value + ' '
		}
	  if (sel == "3")	
	    {
  	    var txt = '<%= l_campo %> = ' + document.datos.texto.value + ' '
		}
      window.returnValue = txt; 
	  window.close();  
    }
}
window.resizeTo(300,180)
</script>

<body leftmargin="0" rightmargin="0" topmargin="0" bottommargin="0">
<form name="datos" method="post" action="#">
<table cellspacing="1" cellpadding="0" border="0" width="100%">
  <tr>
    <td class="th2" colspan="3">Filtrar</td>
  </tr>
<tr>
    <td align="right"><b>Mayor a:</b></td>
	<td><input type="Radio" name="orden" value="1" checked></td>
</tr>
<tr>
    <td align="right"><b>Menor a:</b></td>
	<td><input type="Radio" name="orden" value="2"></td>
</tr>
<tr>
    <td align="right"><b>Igual a:</b></td>
	<td><input type="Radio" name="orden" value="3"></td>
</tr>
<tr>
    <td align="right"><b>Texto:</b></td>
	<td><input type="Text" size="20" maxlength="15" name="texto" value=""></td>
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

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
<title><%= Session("Titulo")%>Filtrar - Ticket</title>
</head>
<script src="/turnos/shared/js/fn_fechas.js"></script>
<script>
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
  	    var txt = '<%= l_fuente %>&orden=<%= l_orden %>&filtro=<%= l_campo %> > ' + document.datos.texto.value + ' '
		}
	  if (sel == "2")	
	    {
  	    var txt = '<%= l_fuente %>&orden=<%= l_orden %>&filtro=<%= l_campo %> < ' + document.datos.texto.value + ' '
		}
	  if (sel == "3")	
	    {
  	    var txt = '<%= l_fuente %>&orden=<%= l_orden %>&filtro=<%= l_campo %> = ' + document.datos.texto.value + ' '
		}
      window.opener.ifrm.location = txt; 
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
		<a class=sidebtnABM href="Javascript:filtrar()">Aceptar</a>
		<a class=sidebtnABM href="Javascript:window.close()">Cancelar</a>		
	</td>
</tr>
</table>
</form>
</body>
</html>

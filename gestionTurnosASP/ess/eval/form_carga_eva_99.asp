<% Option Explicit %>
<!--#include virtual="/ess/ess/shared/inc/const.inc"-->
<%
'Modificado: Leticia Amadio - 13-10-2005 - Adecuacion a Autogestion
%>
<%
Dim l_etiqueta 
Dim l_campos 
Dim l_tipos 
Dim l_orden 
Dim l_destino 
Dim l_cantidad
Dim l_actual
Dim l_liste(20)
Dim l_listc(20)
Dim l_listt(20)
Dim l_i

Dim l_filtro

l_etiqueta = Request.querystring("etiquetas")
l_campos   = Request.querystring("campos")
l_tipos    = Request.querystring("tipos")
l_orden    = Request.querystring("orden")
l_destino  = Request.querystring("Pagina")

l_cantidad = 0
do while len(l_etiqueta) > 0
  if inStr(l_etiqueta,";") <> 0 then
    l_actual   = left(l_etiqueta, inStr(l_etiqueta,";") - 1)
    l_etiqueta = mid (l_etiqueta, inStr(l_etiqueta,";") + 1)
  else
    l_Actual = l_etiqueta
	l_etiqueta = ""
  end if
  l_cantidad = l_cantidad + 1
  l_liste(l_cantidad) = l_actual
loop

l_cantidad = 0
do while len(l_campos) > 0
  if inStr(l_campos,";") <> 0 then
    l_actual = left(l_campos, inStr(l_campos,";") - 1)
    l_campos = mid (l_campos, inStr(l_campos,";") + 1)
  else
    l_Actual = l_campos
	l_campos = ""
  end if
  l_cantidad = l_cantidad + 1
  l_listc(l_cantidad) = l_actual
loop

l_cantidad = 0
do while len(l_tipos) > 0
  if inStr(l_tipos,";") <> 0 then
    l_actual = left(l_tipos, inStr(l_tipos,";") - 1)
    l_tipos = mid (l_tipos, inStr(l_tipos,";") + 1)
  else
    l_Actual = l_tipos
	l_tipos = ""
  end if
  l_cantidad = l_cantidad + 1
  l_listt(l_cantidad) = l_actual
loop
%>

<html>
<head>
<link href="../<%=c_estilo %>" rel="StyleSheet" type="text/css">
<meta http-equiv="Content-Type" content="text/html; charset=iso-8859-1">
<title>Filtrar - Evaluaci&oacute;n de Desempe&ntilde;o - RHPro &reg;</title>
</head>
<script>
function filtrar()
{
var filtro="";
  for (i=0;i<=<%= l_cantidad %>;i++)
  {
    if (document.datos.orden[i].checked)
	  var sel = document.datos.orden[i].value
  }
  if (sel.substr(0,1) == "R") // Restaurar
    {
      filtro = opener.agregarfiltro();
	  window.close();  
    }
  else
    {
	if (sel.substr(0,1) == "T") // Es un texto 
	{
	  filtro = window.showModalDialog("/ess/ess/shared/asp/dial_filtro_browse_txt.asp?campo=" + sel.substr(1) + "&fuente=<%= l_destino %>&orden=<%= l_orden %>",'','dialogWidth:20;dialogHeight:12.2');
	  if (filtro !== "") {
	  	filtro = filtro + " and "+ opener.agregarfiltro();
	  
	  } else {
	  	filtro = opener.agregarfiltro();
	  }
				
	}  
	else
 	  if (sel.substr(0,1) == "N") // Es numerico 
	  {
  	    filtro = window.showModalDialog("/ess/ess/shared/asp/dial_filtro_browse_num.asp?campo=" + sel.substr(1) + "&fuente=<%= l_destino %>&orden=<%= l_orden %>",'','dialogWidth:20;dialogHeight:12.2');
		if (filtro !== "") {
	  		filtro = filtro + " and "+ opener.agregarfiltro();
	  	} else {
	  		filtro = opener.agregarfiltro();
	  	}
	  } else {	// Es fecha
  	    filtro = window.showModalDialog("/ess/ess/shared/asp/dial_filtro_browse_fec.asp?campo=" + sel.substr(1) + "&fuente=<%= l_destino %>&orden=<%= l_orden %>",'','dialogWidth:20;dialogHeight:12.2');
		if (filtro !== "") {
	  		filtro = filtro + " and "+ opener.agregarfiltro();
	   	} else {
	  		filtro = opener.agregarfiltro();
	  	}
	  }
    }

window.opener.ifrm.location = '<%= l_destino %>?orden=<%= l_orden %>&filtro='+filtro; 
window.close();  
	
}

window.resizeTo(300,<%= l_cantidad * 25 + 100 %>)

</script>

<body leftmargin="0" rightmargin="0" topmargin="0" bottommargin="0">
<form name="datos" method="post">
<table cellspacing="1" cellpadding="0" border="0" width="100%">
  <tr>
    <th class="th2" colspan="3">Filtrar</th>
  </tr>
<%
l_i = 1
do while l_i <= l_cantidad
%>  
<tr>
    <td align="right"><b><%= l_liste(l_i)%></b></td>
	<td><input type="Radio" name="orden" value="<%= l_listt(l_i) & l_listc(l_i)%>" <% If l_i = 1 then%>checked<% End If %>></td>
</tr>
<%
  l_i = l_i + 1
loop
%>  
<tr>
    <td align="right"><b>Restaurar:</b></td>
	<td><input type="Radio" name="orden" value="R"></td>
</tr>
</table>
<table width="100%" border="0" cellspacing="0" cellpadding="0">
<tr>
    <td align="right" class="th2">
		<br>
		<a class=sidebtnABM href="Javascript:filtrar()">Aceptar</a>
		<a class=sidebtnABM href="Javascript:window.close()">Cancelar</a>		
	</td>
</tr>
</table>
</form>
</body>
</html>

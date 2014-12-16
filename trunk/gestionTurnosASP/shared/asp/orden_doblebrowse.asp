<% Option Explicit %>
<!--
Archivo: orden_doblebrowse.asp
Descripción: Ordena los item en el doble browse
Autor: F. Favre  
Fecha: 10-03
Modificado:
-->
<%
 Dim l_etiquetas 
 Dim l_funciones
 Dim l_lado
 Dim l_cantidad
 Dim l_actual
 Dim l_listaetiquetas(20)
 Dim l_listaFunciones(20)
 Dim l_i
 
 l_etiquetas	= Request.querystring("Etiquetas")
 l_funciones	= Request.querystring("Funciones")
 l_lado			= Request.querystring("lado")
 
 l_cantidad = 0
 
 do while len(l_etiquetas) > 0
 	if inStr(l_etiquetas,";") <> 0 then
    	l_actual = left(l_etiquetas, inStr(l_etiquetas,";") - 1)
	    l_etiquetas  = mid (l_etiquetas, inStr(l_etiquetas,";") + 1)
  	else
    	l_Actual = l_etiquetas
		l_etiquetas = ""
  	end if
  	l_cantidad = l_cantidad + 1
  	l_listaetiquetas(l_cantidad) = l_actual
 loop
 
 l_cantidad = 0
 
 do while len(l_funciones) > 0
 	if inStr(l_funciones,";") <> 0 then
    	l_actual     = left(l_funciones, inStr(l_funciones,";") - 1)
	    l_funciones  = mid (l_funciones, inStr(l_funciones,";") + 1)
  	else
    	l_Actual    = l_funciones
		l_funciones = ""
  	end if
  	l_cantidad = l_cantidad + 1
  	l_listaFunciones(l_cantidad) = l_actual
 loop
 
%>

<html>
<head>
<link href="/serviciolocal/shared/css/tables3.css" rel="StyleSheet" type="text/css">
<meta http-equiv="Content-Type" content="text/html; charset=iso-8859-1">
<title><%= Session("Titulo")%>Orden - RHPro &reg;</title>
</head>
<script>
var Asc = true;

function valor(){
	for (i=0;i<=<%= l_cantidad * 2 - 1 %>;i++){
    	if (document.datos.orden[i].checked)
	  		var sel = document.datos.orden[i].value
    }
 	if (sel.substr(0,1) == "-"){
		Asc = false; 
		sel = sel.substr(1,sel.length-1)
	}
  	return sel
}
window.resizeTo(360,<%= l_cantidad * 25 + 100 %>)
</script>

<body leftmargin="0" rightmargin="0" topmargin="0" bottommargin="0">

<form name="datos" method="post">
<table cellspacing="1" cellpadding="0" border="0" width="100%" height="100%">
 <tr>
	<td class="th2" colspan="3">Orden</td>
 </tr>
 <tr>
    <td class="th2">&nbsp;</td>
    <td class="th2">Ascendente</td>
    <td class="th2">Descendente</td>
 </tr>
<%
 l_i = 1
 do while l_i <= l_cantidad
%>
 <tr>
    <td align="right"><b><%= l_listaetiquetas(l_i)%></b></td>
	<td><input type="Radio" name="orden" value="<%= l_listafunciones(l_i) %>" <% If l_i = 1 then%>checked<% End If %>></td>
	<td><input type="Radio" name="orden" value="-<%= l_listafunciones(l_i) %>"></td>
 </tr>
<%
 l_i = l_i + 1
loop
%>  
 <tr>
    <td colspan="3" align="right" class="th2">
		<a class=sidebtnABM href="Javascript:window.opener.Orden(valor(), Asc, <%= l_lado%>);window.close();">Aceptar</a>
		<a class=sidebtnABM href="Javascript:window.close()">Cancelar</a>		
	</td>
 </tr>
</table>
</form>
</body>
</html>

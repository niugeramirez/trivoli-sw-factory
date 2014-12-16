<% Option Explicit %>
<!--
Archivo: filtro_doblebrowse_num.asp
Descripción: Filtra items en el doble browse para tipo numericos
Autor: F. Favre  
Fecha: 10-03
Modificado:
-->
<%
 Dim l_campo 
 Dim l_lado
 
 l_campo = request.querystring("campo")
 l_lado  = request.querystring("lado")
%>
<html>
<head>
<link href="/serviciolocal/shared/css/tables3.css" rel="StyleSheet" type="text/css">
<meta http-equiv="Content-Type" content="text/html; charset=iso-8859-1">
<title><%= Session("Titulo")%>Filtrar - RHPro &reg;</title>
</head>
<script src="/serviciolocal/shared/js/fn_numeros.js"></script>
<script>
function filtrar(){
	document.datos.texto.value = document.datos.texto2.value.replace(",", ".");
	for (i=0;i<=2;i++){
		if (document.datos.orden[i].checked)
			var sel = document.datos.orden[i].value
  	}
  	if (document.datos.texto.value == "")
    	alert("Debe ingresar un valor")
  	else
    if (isNaN(document.datos.texto.value)) 
	 	alert("El valor debe ser numérico.");
    else
	if (!validanumero(document.datos.texto,14, 4))
		  alert("El valor permite 14 enteros y 4 decimales como máximo.")
	else{
		if (sel == "1")	
	  	    var txt = '<%= l_campo %> > ' + document.datos.texto.value;
		if (sel == "2")	
	  	    var txt = '<%= l_campo %> < ' + document.datos.texto.value;
	  	if (sel == "3")	
	  	    var txt = '<%= l_campo %> == ' + document.datos.texto.value;
      	window.opener.Filtrar(<%= l_lado %>,txt); 
	  	window.close();  
    }
}
window.resizeTo(300,180)
</script>

<body leftmargin="0" rightmargin="0" topmargin="0" bottommargin="0" onload="JavaScript: document.datos.texto2.focus();">
<form name="datos" method="post">
<input type="hidden" name="texto" value="">
<table cellspacing="1" cellpadding="0" border="0" width="100%" height="100%">
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
		<td><input type="Text" name="texto2" value=""></td>
	</tr>
	<tr>
    	<td colspan="2" align="right" class="th2">
		<a class=sidebtnABM href="Javascript:filtrar()">Aceptar</a>
		<a class=sidebtnABM href="Javascript:window.close()">Cancelar</a>		
		</td>
	</tr>
</table>
</form>
</body>
</html>

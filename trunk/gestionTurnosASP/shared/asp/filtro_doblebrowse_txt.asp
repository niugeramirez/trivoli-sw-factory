<% Option Explicit %>
<!--
Archivo: filtro_doblebrowse_txt.asp
Descripción: Filtra items en el doble browse para tipo Texto
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
<script src="/serviciolocal/shared/js/fn_valida.js"></script>
<script>
function filtrar(){
 var filtro;
	for (i=0;i<=2;i++){
	    if (document.datos.orden[i].checked)
			var sel = document.datos.orden[i].value
	}
  	if (document.datos.texto.value == ""){
      alert("Debe ingresar un texto.");
	  document.datos.texto.focus();
	  return;
    }
    if (!stringValido(document.datos.texto.value)){
	  alert("El Texto contiene caracteres no válidos.");
	  document.datos.texto.select();
	  document.datos.texto.focus();
	  return;
    }
	if (sel == "1")	
  	    var filtro = "<%= l_campo %>.toUpperCase().indexOf('" + document.datos.texto.value.toUpperCase() +"')== 0"
	if (sel == "2")	
  	    var filtro = "<%= l_campo %>.toUpperCase().indexOf('" + document.datos.texto.value.toUpperCase() +"') >= 0" 
	if (sel == "3")	
  	    var filtro = "(<%= l_campo %>.toUpperCase() == '" + document.datos.texto.value.toUpperCase() + "')"
	
   	window.opener.Filtrar(<%= l_lado %>, filtro);
  	window.close();  
    
}
window.resizeTo(300,180)
</script>

<body leftmargin="0" rightmargin="0" topmargin="0" bottommargin="0" onload="JavaScript: document.datos.texto.focus();">
<form name="datos" method="post" action="#">
<table cellspacing="1" cellpadding="0" border="0" width="100%" height="100%">
	<tr>
    	<td class="th2" colspan="3">Filtrar</td>
	</tr>
	<tr>
	    <td align="right"><b>Comienza con:</b></td>
		<td><input type="Radio" name="orden" value="1" checked></td>
	</tr>
	<tr>
	    <td align="right"><b>Contiene:</b></td>
		<td><input type="Radio" name="orden" value="2"></td>
	</tr>
	<tr>
	    <td align="right"><b>Igual a:</b></td>
		<td><input type="Radio" name="orden" value="3"></td>
	</tr>
	<tr>
	    <td align="right"><b>Texto:</b></td>
		<td><input type="Text" name="texto" value=""></td>
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

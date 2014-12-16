<% Option Explicit %>
<!--
Archivo: filtro_doblebrowse_bool.asp
Descripción: Filtra items en el doble browse para tipo booleanos
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
<link href="/turnos/shared/css/tables3.css" rel="StyleSheet" type="text/css">
<meta http-equiv="Content-Type" content="text/html; charset=iso-8859-1">
<title><%= Session("Titulo")%>Filtrar - RHPro &reg;</title>
</head>
<script>
function filtrar(){
	if (document.datos.orden[0].checked)	
  	    var txt = '<%= l_campo %> == -1';
	else
  	    var txt = '<%= l_campo %> == 0';
   	window.opener.Filtrar(<%= l_lado %>,txt); 
  	window.close();  
}
window.resizeTo(250,120)
</script>

<body leftmargin="0" rightmargin="0" topmargin="0" bottommargin="0">
<form name="datos" method="post">
<table cellspacing="1" cellpadding="0" border="0" width="100%" height="100%">
	<tr>
    	<td class="th2" colspan="3">Filtrar</td>
  	</tr>
	<tr>
    	<td align="right"><b>S&iacute;:</b></td>
		<td><input type="Radio" name="orden" value="1" checked></td>
	</tr>
	<tr>
	    <td align="right"><b>No:</b></td>
		<td><input type="Radio" name="orden" value="2"></td>
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

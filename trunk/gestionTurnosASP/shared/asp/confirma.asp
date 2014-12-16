<% Option Explicit %>
<!--#include virtual="/turnos/shared/db/conn_db.inc"-->

<% 
 Dim l_texto
 Dim l_foco
 
 l_texto 	 = request.querystring("texto")
 l_foco 	 = request.querystring("foco")
%>
<html>
<head>
<link href="/turnos/shared/css/tables3.css" rel="StyleSheet" type="text/css">
<meta http-equiv="Content-Type" content="text/html; charset=iso-8859-1">
<title><%= Session("Titulo")%>Confirma - RHPro &reg;</title>
</head>
<script src="/turnos/shared/js/fn_ayuda.js"></script>
<script>
window.returnValue=false;
function focalizar(opcion){
if (opcion == 'cancelar')
	document.all.cancelar.focus();
else
	document.all.aceptar.focus();
}
</script>
<body leftmargin="0" rightmargin="0" topmargin="0" bottommargin="0" onload="JavaScript:focalizar('<%= l_foco %>')">
<table cellspacing="0" cellpadding="0" border="0" width="100%" height="100%">
 <tr>
    <td class="th2" height="1">Confirma </td>
	<td align="right" class="barra" valign="middle">
		<a class=sidebtnHLP href="#" onclick="Javascript:ayuda('<%= Request.ServerVariables("SCRIPT_NAME")%>');">Ayuda</a>
	</td>
 </tr>
 <tr>
   	<td align="center" colspan="2"><%= l_texto %></td>
 </tr>


 <tr>
	<td align="right" class="th2" colspan="4" height="1">
		<a name="aceptar" class=sidebtnABM onclick="Javascript:window.returnValue=true;window.close();" onkeypress="Javascript:window.returnValue=true;window.close();" href="#">Aceptar</a>
		<a name="cancelar" class=sidebtnABM onclick="Javascript:window.returnValue=false;window.close();" onkeypress="Javascript:window.returnValue=false;window.close();" href="#">Cancelar</a>
	</td>
 </tr>
</table>
</body>
</html>

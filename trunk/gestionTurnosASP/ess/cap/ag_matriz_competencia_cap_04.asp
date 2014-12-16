<% Option Explicit %>
<!--#include virtual="/serviciolocal/shared/db/conn_db.inc"-->
<!--#include virtual="/serviciolocal/shared/inc/fecha.inc"-->
<!--
Archivo: ag_matriz_competencia_cap_04.asp
Descripción: 
Autor : Raul Chinestra

-->
<%
dim l_rs
dim l_sql
dim l_titulo
dim l_tenro

dim l_tipo
dim l_valor

dim l_total

l_tipo	 = Request.QueryString("radio")
l_valor  = Request.QueryString("porcen")

%>
<html>
<head>
<link href="../<%= session("estilo")%>" rel="StyleSheet" type="text/css">
<meta http-equiv="Content-Type" content="text/html; charset=iso-8859-1">
<title>Ingrese el Porcentaje - RHPro &reg;</title>
</head>
<script src="/serviciolocal/shared/js/fn_windows.js"></script>
<script src="/serviciolocal/shared/js/fn_confirm.js"></script>
<script src="/serviciolocal/shared/js/fn_ayuda.js"></script>
<script>
function Validar_Formulario(){

	if ((document.datos.valor.value > 100)||(document.datos.valor.value < 1)){
		alert("Debe Ingresar un Porcentaje entre 1 y 100.");document.datos.valor.focus();
	}else{
		   if (isNaN(document.datos.valor.value))
		      { alert("Debe Ingresar un Porcentaje Numérico.");
		        document.datos.valor.focus();
		      } else{ window.returnValue = document.datos.valor.value;
			 	      window.close();
			        }
			
		 }
	
}

function chek(valor){
	if (valor == 1){
		document.datos.valorsum.checked = true;
		document.datos.valormax.checked = false;
		document.datos.valor.value = document.datos.valorsum.value;
	}else{
		document.datos.valorsum.checked = false;
		document.datos.valormax.checked = true;
		document.datos.valor.value = document.datos.valormax.value;
	}
 }
</script>
<body  leftmargin="0" rightmargin="-1" topmargin="0" bottommargin="0" onload="document.datos.valor.focus();">
<form name="datos" method="post" >
<table cellspacing="0" cellpadding="0" border="0" width="100%" height="100%">
 <tr>
    <td class="th2">
		<!--Valor del <%'= l_titulo %>-->
	</td>
	<td align="right" class="barra" >
		<a class=sidebtnHLP onclick="Javascript:ayuda('<%= Request.ServerVariables("SCRIPT_NAME")%>');">Ayuda</a>
	</td>
</tr>
<tr>
	<td align="center" colspan="2">
	<b>Porcentaje:</b>
	<input type="textbox" name=valor value="<%= l_valor %>" maxlength="3" size="3" onkeypress="">
	<b>%</b>
	</td>
</tr>
<tr>
    <td colspan="2" align="right" class="th2">
    <a class=sidebtnABM onclick="Javascript:Validar_Formulario()">Aceptar</a>
	<a class=sidebtnABM onclick="Javascript:window.close()">Cancelar</a>
	</td>
</tr>
</table>
</form>
</body>
</html>


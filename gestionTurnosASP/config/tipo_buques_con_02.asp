<% Option Explicit %>
<!--#include virtual="/serviciolocal/shared/db/conn_db.inc"-->
<% 
'Archivo: companies_con_02.asp
'Descripción: ABM de Companies
'Autor : Raul Chinestra
'Fecha: 26/11/2007

'Datos del formulario
Dim l_tipbuqnro
Dim l_tipbuqdes

'ADO
Dim l_tipo
Dim l_sql
Dim l_rs

l_tipo = request.querystring("tipo")

%>
<html>
<head>
<link href="/serviciolocal/shared/css/tables3.css" rel="StyleSheet" type="text/css">
<meta http-equiv="Content-Type" content="text/html; charset=iso-8859-1">
<title><%= Session("Titulo")%> Tipos de Buques - Buques</title>
</head>
<script src="/serviciolocal/shared/js/fn_ayuda.js"></script>
<script src="/serviciolocal/shared/js/fn_windows.js"></script>
<script src="/serviciolocal/shared/js/fn_valida.js"></script>
<script>
function Validar_Formulario(){

if (Trim(document.datos.tipbuqdes.value) == ""){
	alert("Debe ingresar la Descripción.");
	document.datos.tipbuqdes.focus();
	return;
}

if (!stringValido(document.datos.tipbuqdes.value)){
	alert("La Descripción contiene caracteres inválidos.");
	document.datos.tipbuqdes.focus();
	return;
}

var d=document.datos;
document.valida.location = "tipo_buques_con_06.asp?tipo=<%= l_tipo%>&tipbuqnro="+document.datos.tipbuqnro.value + "&tipbuqdes="+document.datos.tipbuqdes.value;

}

function valido(){
	document.datos.submit();
}

function invalido(texto){
	alert(texto);
	document.datos.tipbuqdes.focus();
}

</script>
<% 
select Case l_tipo
	Case "A":
		l_tipbuqdes = ""
	Case "M":
		Set l_rs = Server.CreateObject("ADODB.RecordSet")
		l_tipbuqnro = request.querystring("cabnro")
		l_sql = "SELECT * "
		l_sql = l_sql & " FROM buq_tipobuque "
		l_sql  = l_sql  & " WHERE tipbuqnro =" & l_tipbuqnro
		rsOpen l_rs, cn, l_sql, 0 
		if not l_rs.eof then
			l_tipbuqdes = l_rs("tipbuqdes")
		end if
		l_rs.Close
end select
%>
<body leftmargin="0" rightmargin="0" topmargin="0" bottommargin="0" onload="JavaScript:document.datos.tipbuqdes.focus()">
<form name="datos" action="tipo_buques_con_03.asp?tipo=<%= l_tipo %>" method="post" target="valida">
<input type="Hidden" name="tipbuqnro" value="<%= l_tipbuqnro %>">


<table cellspacing="0" cellpadding="0" border="0" width="100%" height="100%">
<tr>
    <td class="th2" nowrap>Tipos de Buques</td>
	<td class="th2" align="right">
		<!--
		<a class=sidebtnHLP href="Javascript:ayuda('<%'= Request.ServerVariables("SCRIPT_NAME")%>');">Ayuda</a>
		-->
	</td>
</tr>
<tr>
	<td colspan="2" height="100%">
		<table border="0" cellspacing="0" cellpadding="0">
			<tr>
				<td width="50%"></td>
				<td>
					<table cellspacing="0" cellpadding="0" border="0">
					<tr>
					    <td align="right"><b>Descripción:</b></td>
						<td>
							<input type="text" name="tipbuqdes" size="35" maxlength="25" value="<%= l_tipbuqdes %>">
						</td>
					</tr>			
					</table>
				</td>
				<td width="50%"></td>
			</tr>
		</table>
	</td>
</tr>
<tr>
    <td colspan="2" align="right" class="th2">
		<a class=sidebtnABM href="Javascript:Validar_Formulario()">Aceptar</a>
		<a class=sidebtnABM href="Javascript:window.close()">Cancelar</a>
	</td>
</tr>
</table>
<iframe name="valida" style="visibility=hidden;" src="" width="100%" height="100%"></iframe> 
</form>
<%
set l_rs = nothing
cn.Close
set cn = nothing
%>
</body>
</html>

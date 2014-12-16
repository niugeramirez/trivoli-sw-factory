<% Option Explicit %>
<!--#include virtual="/serviciolocal/shared/db/conn_db.inc"-->
<% 
'Archivo: companies_con_02.asp
'Descripción: ABM de Companies
'Autor : Raul Chinestra
'Fecha: 26/11/2007

'Datos del formulario
Dim l_expnro
Dim l_expdes

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
<title><%= Session("Titulo")%> Exportadoras - Buques</title>
</head>
<script src="/serviciolocal/shared/js/fn_ayuda.js"></script>
<script src="/serviciolocal/shared/js/fn_windows.js"></script>
<script src="/serviciolocal/shared/js/fn_valida.js"></script>
<script>
function Validar_Formulario(){

if (Trim(document.datos.expdes.value) == ""){
	alert("Debe ingresar la Descripción.");
	document.datos.expdes.focus();
	return;
}

if (!stringValido(document.datos.expdes.value)){
	alert("La Descripción contiene caracteres inválidos.");
	document.datos.expdes.focus();
	return;
}

var d=document.datos;
document.valida.location = "exportadora_con_06.asp?tipo=<%= l_tipo%>&expnro="+document.datos.expnro.value + "&expdes="+document.datos.expdes.value;

}

function valido(){
	document.datos.submit();
}

function invalido(texto){
	alert(texto);
	document.datos.expdes.focus();
}

</script>
<% 
select Case l_tipo
	Case "A":
		l_expdes = ""
	Case "M":
		Set l_rs = Server.CreateObject("ADODB.RecordSet")
		l_expnro = request.querystring("cabnro")
		l_sql = "SELECT * "
		l_sql = l_sql & " FROM buq_exportadora "
		l_sql  = l_sql  & " WHERE expnro =" & l_expnro
		rsOpen l_rs, cn, l_sql, 0 
		if not l_rs.eof then
			l_expdes = l_rs("expdes")
		end if
		l_rs.Close
end select
%>
<body leftmargin="0" rightmargin="0" topmargin="0" bottommargin="0" onload="JavaScript:document.datos.expdes.focus()">
<form name="datos" action="exportadora_con_03.asp?tipo=<%= l_tipo %>" method="post" target="valida">
<input type="Hidden" name="expnro" value="<%= l_expnro %>">


<table cellspacing="0" cellpadding="0" border="0" width="100%" height="100%">
<tr>
    <td class="th2" nowrap>Exportadora</td>
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
							<input type="text" name="expdes" size="35" maxlength="25" value="<%= l_expdes %>">
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

<% Option Explicit %>
<!--#include virtual="/serviciolocal/shared/db/conn_db.inc"-->
<% 
'Archivo: companies_con_02.asp
'Descripción: ABM de Companies
'Autor : Raul Chinestra
'Fecha: 26/11/2007

'Datos del formulario
Dim l_mernro
Dim l_merdes
Dim l_tipmerdes
Dim l_merord

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
<title><%= Session("Titulo")%> Mercaderías - Buques</title>
</head>
<script src="/serviciolocal/shared/js/fn_ayuda.js"></script>
<script src="/serviciolocal/shared/js/fn_windows.js"></script>
<script src="/serviciolocal/shared/js/fn_valida.js"></script>
<script>
function Validar_Formulario(){

if (Trim(document.datos.merdes.value) == ""){
	alert("Debe ingresar la Descripción.");
	document.datos.merdes.focus();
	return;
}

if (!stringValido(document.datos.merdes.value)){
	alert("La Descripción contiene caracteres inválidos.");
	document.datos.merdes.focus();
	return;
}

var d=document.datos;
document.valida.location = "mercaderias_con_06.asp?tipo=<%= l_tipo%>&mernro="+document.datos.mernro.value + "&merdes="+document.datos.merdes.value;

}

function valido(){
	document.datos.submit();
}

function invalido(texto){
	alert(texto);
	document.datos.merdes.focus();
}

</script>
<% 
select Case l_tipo
	Case "A":
		l_merdes = ""
		l_tipmerdes = ""
		l_merord = "0"
	Case "M":
		Set l_rs = Server.CreateObject("ADODB.RecordSet")
		l_mernro = request.querystring("cabnro")
		l_sql = "SELECT * "
		l_sql = l_sql & " FROM buq_mercaderia "
		l_sql  = l_sql  & " WHERE mernro =" & l_mernro
		rsOpen l_rs, cn, l_sql, 0 
		if not l_rs.eof then
			l_merdes = l_rs("merdes")
			l_tipmerdes = l_rs("tipmerdes")
			l_merord = l_rs("merord")
		end if
		l_rs.Close
end select
%>
<body leftmargin="0" rightmargin="0" topmargin="0" bottommargin="0" onload="JavaScript:document.datos.merdes.focus()">
<form name="datos" action="mercaderias_con_03.asp?tipo=<%= l_tipo %>" method="post" target="valida">
<input type="Hidden" name="mernro" value="<%= l_mernro %>">


<table cellspacing="0" cellpadding="0" border="0" width="100%" height="100%">
<tr>
    <td class="th2" nowrap>Mercaderias</td>
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
							<input type="text" name="merdes" size="35" maxlength="25" value="<%= l_merdes %>">
						</td>
					</tr>	
					<tr>
					    <td align="right" nowrap><b>Tipo Mercadería:</b></td>					
						<td>
							<select name="tipmerdes" size="1" style="width:100;" >
							<option value="" selected>&nbsp;</option>
								<option value="CAS">CAS</option>
								<option value="INF">INF</option>								
								<option value="OTR">OTR</option>								
								<option value="PC">PC</option>																								
							</select>
							<script> document.datos.tipmerdes.value= "<%= l_tipmerdes %>"</script>												
						</td>
					</tr>	
					<tr>
					    <td align="right"><b>Orden:</b></td>
						<td>
							<input type="text" name="merord" size="5" maxlength="5" value="<%= l_merord %>">
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

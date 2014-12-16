<% Option Explicit %>
<!--#include virtual="/serviciolocal/shared/db/conn_db.inc"-->
<% 
'Archivo: berths_con_02.asp
'Descripción: ABM de berths
'Autor : Raul Chinestra
'Fecha: 23/11/2007

'Datos del formulario
Dim l_connro
Dim l_mernro
Dim l_expnro
Dim l_sitnro
Dim l_conton
Dim l_desnro

'ADO
Dim l_tipo
Dim l_sql
Dim l_rs

Dim l_buqnro

l_tipo   = request.querystring("tipo")
l_buqnro = request("buqnro")

%>
<html>
<head>
<link href="/serviciolocal/shared/css/tables3.css" rel="StyleSheet" type="text/css">
<meta http-equiv="Content-Type" content="text/html; charset=iso-8859-1">
<title><%= Session("Titulo")%> Contenidos - Buques</title>
</head>
<script src="/serviciolocal/shared/js/fn_ayuda.js"></script>
<script src="/serviciolocal/shared/js/fn_windows.js"></script>
<script src="/serviciolocal/shared/js/fn_valida.js"></script>
<script>
function Validar_Formulario(){


if (Trim(document.datos.mernro.value) == "0"){
	alert("Debe ingresar la Mercadería.");
	document.datos.mernro.focus();
	return;
}
/*
if (!stringValido(document.datos.berdes.value)){
	alert("La Descripción contiene caracteres inválidos.");
	document.datos.berdes.focus();
	return;
}
var d=document.datos;
document.valida.location = "berths_con_06.asp?tipo=<%= l_tipo%>&bernro="+document.datos.bernro.value + "&berdes="+document.datos.berdes.value;
*/

valido();
}

function valido(){
	document.datos.submit();
}

function invalido(texto){
	alert(texto);
	document.datos.berdes.focus();
}

</script>
<% 
select Case l_tipo
	Case "A":
		l_mernro = 0
		l_expnro = 0
		l_sitnro = 0
		l_conton = 0
		l_desnro = 0
		
	Case "M":
		Set l_rs = Server.CreateObject("ADODB.RecordSet")
		l_connro = request.querystring("cabnro")
		l_sql = "SELECT  mernro, expnro, sitnro, conton, desnro "
		l_sql = l_sql & " FROM buq_contenido "
		l_sql  = l_sql  & " WHERE buqnro = " & l_buqnro
		rsOpen l_rs, cn, l_sql, 0 
		if not l_rs.eof then
			l_mernro = l_rs("mernro")
			l_expnro = l_rs("expnro")
			l_sitnro = l_rs("sitnro")
			l_conton = l_rs("conton")
			l_desnro = l_rs("desnro")
		end if
		l_rs.Close
end select
%>
<body leftmargin="0" rightmargin="0" topmargin="0" bottommargin="0" onload="JavaScript:document.datos.mernro.focus()">
<form name="datos" action="contenidos_con_03.asp?tipo=<%= l_tipo %>" method="post" target="valida">
<input type="Hidden" name="buqnro" value="<%= l_buqnro %>">
<input type="Hidden" name="connro" value="<%= l_connro %>">


<table cellspacing="0" cellpadding="0" border="0" width="100%" height="100%">
<tr>
    <td class="th2" nowrap>Contenidos</td>
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
						<td align="right"><b>Mercadería:</b></td>
						<td><select name="mernro" size="1" style="width:250;">
								<option value=0 selected>&nbsp;</option>
								<%Set l_rs = Server.CreateObject("ADODB.RecordSet")
								l_sql = "SELECT  * "
								l_sql  = l_sql  & " FROM buq_mercaderia "
								l_sql  = l_sql  & " ORDER BY merdes "
								rsOpen l_rs, cn, l_sql, 0
								do until l_rs.eof		%>	
								<option value= <%= l_rs("mernro") %> > 
								<%= l_rs("merdes") %> (<%=l_rs("mernro")%>) </option>
								<%	l_rs.Movenext
								loop
								l_rs.Close %>
							</select>
							<script>document.datos.mernro.value= "<%= l_mernro %>"</script>
						</td>	
					</tr>
					<tr>
						<td align="right"><b>Exportadora:</b></td>
						<td><select name="expnro" size="1" style="width:250;">
								<option value=0 selected>&nbsp;</option>
								<%Set l_rs = Server.CreateObject("ADODB.RecordSet")
								l_sql = "SELECT  * "
								l_sql  = l_sql  & " FROM buq_exportadora "
								l_sql  = l_sql  & " ORDER BY expdes "
								rsOpen l_rs, cn, l_sql, 0
								do until l_rs.eof		%>	
								<option value= <%= l_rs("expnro") %> > 
								<%= l_rs("expdes") %> (<%=l_rs("expnro")%>) </option>
								<%	l_rs.Movenext
								loop
								l_rs.Close %>
							</select>
							<script>document.datos.expnro.value= "<%= l_expnro %>"</script>
						</td>	
						</tr>						
					<tr>
						<td align="right"><b>Sitio:</b></td>
						<td><select name="sitnro" size="1" style="width:250;">
								<option value=0 selected>&nbsp;</option>
								<%Set l_rs = Server.CreateObject("ADODB.RecordSet")
								l_sql = "SELECT  * "
								l_sql  = l_sql  & " FROM buq_sitio "
								l_sql  = l_sql  & " ORDER BY sitdes "
								rsOpen l_rs, cn, l_sql, 0
								do until l_rs.eof		%>	
								<option value= <%= l_rs("sitnro") %> > 
								<%= l_rs("sitdes") %> (<%=l_rs("sitnro")%>) </option>
								<%	l_rs.Movenext
								loop
								l_rs.Close %>
							</select>
							<script>document.datos.sitnro.value= "<%= l_sitnro %>"</script>
							</td>	
						</tr>						
						<tr>				
						    <td align="right"><b>Toneladas:</b></td>
							<td>
								<input type="text" name="conton" size="10" maxlength="10" value="<%= l_conton %>">
							</td>
						</tr>
					<tr>
						<td align="right"><b>Destino:</b></td>
						<td><select name="desnro" size="1" style="width:250;">
								<option value=0 selected>&nbsp;</option>
								<%Set l_rs = Server.CreateObject("ADODB.RecordSet")
								l_sql = "SELECT  * "
								l_sql  = l_sql  & " FROM buq_destino "
								l_sql  = l_sql  & " ORDER BY desdes "
								rsOpen l_rs, cn, l_sql, 0
								do until l_rs.eof		%>	
								<option value= <%= l_rs("desnro") %> > 
								<%= l_rs("desdes") %> (<%=l_rs("desnro")%>) </option>
								<%	l_rs.Movenext
								loop
								l_rs.Close %>
							</select>
							<script>document.datos.desnro.value= "<%= l_desnro %>"</script>
							</td>	
						</tr>													
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

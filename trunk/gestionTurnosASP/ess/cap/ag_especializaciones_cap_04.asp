<% Option Explicit %>
<!--#include virtual="/serviciolocal/ess/shared/inc/const.inc"-->
<!--#include virtual="/serviciolocal/shared/db/conn_db.inc"-->
<!--#include virtual="/serviciolocal/shared/inc/fecha.inc"-->
<!--#include virtual="/serviciolocal/ess/shared/inc/accesoESS.inc"-->
<!--
Archivo: ag_especializaciones_cap_04.asp
Descripcion: especializaciones
Autor: Lisandro Moro
Fecha: 29/03/2004
Modificado:
-->
<%
on error goto 0

dim l_rs
dim l_sql
Dim l_nivel

l_nivel	=  Request.QueryString("nivel")
%>
<html>
<head>
<link href="../<%= c_estilo %>" rel="StyleSheet" type="text/css">
<meta http-equiv="Content-Type" content="text/html; charset=iso-8859-1">
<title>Ingrese el Porcentaje - RHPro &reg;</title>
</head>
<script src="/serviciolocal/shared/js/fn_windows.js"></script>
<script src="/serviciolocal/shared/js/fn_confirm.js"></script>
<script src="/serviciolocal/shared/js/fn_ayuda.js"></script>
<script>
function Validar_Formulario(){
	var desc;
	if (document.datos.espnro.value == 0){
		alert("Debe seleccionar un nivel");
	}else{
		//alert(document.datos.valor.value);
		switch (document.datos.espnro.value){
		<%Set l_rs = Server.CreateObject("ADODB.RecordSet")
		l_sql = "SELECT espnivnro, espnivdesabr "  
		l_sql = l_sql & " FROM espnivel "
		rsOpen l_rs, cn, l_sql, 0
		l_rs.MoveFirst
		do until l_rs.eof		%>	
			case "<%= l_rs("espnivnro") %>":
				desc = "<%= l_rs("espnivdesabr") %>";
				break;
		<% l_rs.Movenext
		loop
		l_rs.Close%>
		}
		window.returnValue = desc + ";" + document.datos.espnro.value;
		window.close();
	}
}

</script>
<body  leftmargin="0" rightmargin="-1" topmargin="0" bottommargin="0" onload="document.datos.espnro.focus();">
<form name="datos" method="post" >
<table cellspacing="0" cellpadding="0" border="0" width="100%" height="100%">
 <tr>
    <td class="th2">
		<!--Valor del <%'= l_titulo %>-->
	</td>
	<td align="right" class="th2" >
		&nbsp;
	</td>
</tr>
<tr>
	<td align="center" colspan="2">
	<b>Nivel:</b>
		<%Set l_rs = Server.CreateObject("ADODB.RecordSet")
		l_sql = "SELECT espnivnro, espnivdesabr "  
		l_sql = l_sql & " FROM espnivel "
		rsOpen l_rs, cn, l_sql, 0 %>
		<select name=espnro size="1">
		<option value=0 selected><< Seleccione una Opción >></option>
		<%	l_rs.MoveFirst
		do until l_rs.eof		%>	
			<option value=<%= l_rs("espnivnro") %> > 
			<%= l_rs("espnivdesabr") %> (<%=l_rs("espnivnro")%>) </option>
			<% l_rs.Movenext
		loop
		l_rs.Close %>	
	</select>
	<% If l_nivel <> "" then %>
		<script>document.datos.espnro.value = <%= l_nivel %></script>
	<% End If %>
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


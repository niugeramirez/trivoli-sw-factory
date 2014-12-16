<% Option Explicit %>
<!--#include virtual="/turnos/shared/inc/sec.inc"-->
<!--#include virtual="/turnos/shared/inc/const.inc"-->
<!--#include virtual="/turnos/shared/db/conn_db.inc"-->
<% 
Dim rs
Dim sql

Set rs = Server.CreateObject("ADODB.RecordSet")
%>
<html>
<head>
<link href="/turnos/shared/css/tables3.css" rel="StyleSheet" type="text/css">
<title><%= Session("Titulo")%>Acceso por Men&uacute; - Ticket</title>
<script src="/turnos/shared/js/fn_windows.js"></script>
<script src="/turnos/shared/js/fn_confirm.js"></script>
<script src="/turnos/shared/js/fn_ayuda.js"></script>
<script src="/turnos/shared/js/fn_ay_generica.js"></script>
<script>
function actualizar(){
	if (document.datos.menuraiz.value != "-1")
		document.menu1.location = "acceso_menu_01.asp?menuraiz=" + document.datos.menuraiz.value
	else
		document.menu1.location = "blanc.html";
	document.menu2.location = "blanc.html";
}

function modificar(){
	if (document.menu1.jsSelRow == null)
		alert("Debe seleccionar un item.")
	else{
		var donde;
		donde = 'acceso_menu_03.asp?menuaccess=' + document.menu2.datos.menuaccess.value;
		donde += '&menuimg=' + document.menu2.datos.menuimg.value;
		donde += '&menuorder=' + document.menu2.datos.menuorder.value;
		donde += '&menuraiz=' + document.menu2.datos.menuraiz.value;
		abrirVentanaH(donde,'',150,150);
	}
}
</script>

</head>
<body leftmargin="0" topmargin="0" rightmargin="0" bottommargin="0">
<form name="datos">
<table border="0" cellpadding="0" cellspacing="0" width="100%" height="100%">
	<tr style="border-color :CadetBlue;">
		<td align="left" class="barra">Acceso por Men&uacute;</td>
		<td align="right" class="barra">
			<a class=sidebtnHLP href="Javascript:ayuda('<%= Request.ServerVariables("SCRIPT_NAME")%>');">Ayuda</a>
		</td>
	</tr>
	<tr>
		<td colspan="2">
			<b>Ra&iacute;z:</b>
			<% 
			sql = "SELECT menunro, menudesc, menudescext FROM menuraiz order by menudesc"
			rsOpen rs, cn, sql, 0
			%>		     
			<select name="menuraiz" onchange="Javascript:actualizar()">
			<option value="-1" SELECTED>Ninguno</option>
			<% do until rs.eof %>
			<option value="<%= rs("menunro") %>"><%= rs("menudescext") %></option>
			<%
			rs.MoveNext
			loop
			rs.Close
			%>
			</select>
		</td>
	</tr>
	<tr>
		<td colspan="2">
			<b>Menu:</b>
		</td>
	</tr>
	<tr>
		<td height="100%" colspan="2">
			<iframe name="menu1" src="blanc.html" width="100%" height="100%"></iframe> 
		</td>
	</tr>
	<tr>
		<td height="130" colspan="2">
			<b>Datos a Modificar:</b><br>
			<iframe name="menu2" src="blanc.html" width="100%" height="120" scrolling="No"></iframe> 
		</td>
	</tr>
</table>
</form>
</body>
</html>

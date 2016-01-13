<% Option Explicit %>
<!--#include virtual="/turnos/shared/db/conn_db.inc"-->
<% 
Dim l_mulnro
Dim l_mulnom
Dim l_multiple
Dim rs
Dim sql
Dim tipo
%>
<% 
tipo = request("tipo")
%>
<html>
<head>
<link href="/turnos/shared/css/tables3.css" rel="StyleSheet" type="text/css">
<meta http-equiv="Content-Type" content="text/html; charset=iso-8859-1">
<title><%= Session("Titulo")%>Configuración Multiempresa - Ticket</title>
</head>
<script>
function Validar_Formulario()
{
if (document.datos.mulnom.value == "") 
	alert("Debe ingresar la descripcion.");
else
	{
	document.datos.submit();
	}
}

function Nuevo_Dialogo(w_in, pagina, ancho, alto)
{
 return w_in.showModalDialog(pagina,'', 'center:yes;dialogWidth:' + ancho.toString() + ';dialogHeight:' + alto.toString() + ';');
}
function Ayuda_Fecha(txt)
{
 var jsFecha = Nuevo_Dialogo(window, '/turnos/shared/js/calendar.html', 16, 15);

 if (jsFecha == null) txt.value = ''
 else txt.value = jsFecha;
}
</script>
<% 
select Case tipo
	Case "A":
		l_mulnom = ""
		l_multiple = 0
	Case "M":
		l_mulnro = request.QueryString("mulnro")
		Set rs = Server.CreateObject("ADODB.RecordSet")
		sql = "SELECT mulnro, mulnom, multiple "
		sql = sql & "FROM multiempresa "
		sql = sql & "WHERE mulnro = " & l_mulnro
		rsOpen rs, cn, sql, 0 
		if not rs.eof then
			l_mulnro = rs("mulnro")
			l_mulnom = rs("mulnom")
			l_multiple = rs("multiple")
		end if
		rs.Close
		set rs = nothing
end select

%>
<body leftmargin="0" rightmargin="0" topmargin="0" bottommargin="0" onload="document.datos.mulnom.focus()">
<form name="datos" action="multiempresa_gti_03.asp?Tipo=<%= tipo %>" method="post">
<input type="Hidden" name="mulnro" value="<%= l_mulnro %>">
<table cellspacing="1" cellpadding="0" border="0" width="100%">
  <tr>
    <td class="th2" colspan="4">Tablas</td>
  </tr>
<tr>
    <td align="right"><b>Nombre</b></td>
	<td><input type="text" name="mulnom" size="35" maxlength="35" value="<%= l_mulnom %>"></td>
</tr>
<tr>
    <td align="right"><b>Multiempresa</b></td>
	<td>
	<%if l_multiple = 0 then%>
	<input type="checkbox" name="multiple">
	<%else%>
	<input type="checkbox" checked name="multiple">
	<%end if
	%>
	</td>
</tr>
</table>
<table width="100%" border="0" cellspacing="0" cellpadding="0">
<tr>
    <td align="right" class="th2">
		<a class=sidebtnABM href="Javascript:Validar_Formulario()">Aceptar</a>
		<a class=sidebtnABM href="Javascript:window.close()">Cancelar</a>
	</td>
</tr>
</table>
</form>
</body>
</html>

<% Option Explicit %>
<!--#include virtual="/turnos/shared/db/conn_db.inc"-->
<!--
-----------------------------------------------------------------------------
Archivo        : alcance_estructura_02.asp
Descripcion    : 
Modificacion   :
    29/03/2004 - Scarpa D. - Validacion al cargar los perfiles
-----------------------------------------------------------------------------
-->
<% 
' Modificado: 19/09/2003 - CCRossi - que muestre todos los perfiles en la modificacion 
'									 posicionandose en el primero de los elegidos
Dim l_menuraiz
Dim l_menuorder
Dim l_nombre
Dim l_pagina
Dim l_accesos
Dim l_tipo
Dim l_sql
Dim l_rs
Dim l_primero
Dim l_primero1

l_tipo = request("tipo")
l_menuraiz = request("menuraiz")
l_menuorder = request("menuorder")

%>
<html>
<head>
<link href="/turnos/shared/css/tables3.css" rel="StyleSheet" type="text/css">
<meta http-equiv="Content-Type" content="text/html; charset=iso-8859-1">
<title><%= Session("Titulo")%>Botones - Ticket</title>
</head>
<script src="/turnos/shared/js/fn_ayuda.js"></script>
<script>
function Validar_Formulario()
{
if (document.datos.nombre.value == "") 
	alert("Debe ingresar el nombre del boton.");
else
if (document.datos.pagina.value == "") 
	alert("Debe ingresar el nombre de la página.");
else
if (document.datos.menuaccess.value == "") 
	alert("Debe ingresar los perfiles asignados.");
else
	{
	document.datos.submit();
	}
}

function agregar()
{
if (document.datos.menuaccess.value == "*")
	document.datos.menuaccess.value= "";
	
if ((document.datos.menuaccess.value.length + document.datos.perfiles.value.length) < 200){	
	
	if (document.datos.menuaccess.value != "")
		document.datos.menuaccess.value= document.datos.menuaccess.value + ';';
	document.datos.menuaccess.value= document.datos.menuaccess.value + document.datos.perfiles.value
}
	
}
window.resizeTo(500,200);
</script>
<% 

select Case l_tipo
	Case "A":
		l_nombre = ""
		l_pagina = ""
		l_accesos = ""
	Case "M":
		l_nombre = request("cabnro")
		l_pagina = request("pagina")
		Set l_rs = Server.CreateObject("ADODB.RecordSet")		
		l_sql = "SELECT btnaccess "
		l_sql = l_sql & "FROM menubtn "
		l_sql = l_sql & "WHERE menuraiz = " & l_menuraiz & " AND menuorder = " & l_menuorder & " AND btnpagina = '" &  l_pagina & "' AND btnnombre = '" &  l_nombre & "'"
		l_rs.MaxRecords = 1
		rsOpen l_rs, cn, l_sql, 0 
		if not l_rs.eof then
			l_accesos = l_rs("btnaccess")
		end if
		if l_accesos <> "" then
			l_primero = split(l_accesos,";")
			l_primero1 = l_primero(0)
		end if	
		l_rs.Close
		set l_rs =nothing
end select

%>
<body leftmargin="0" rightmargin="0" topmargin="0" bottommargin="0">
<form name="datos" action="armado_menu_13.asp?tipo=<%= l_tipo %>" method="post">
<input type="Hidden" name="menuraiz" value="<%= l_menuraiz %>">
<input type="Hidden" name="menuorder" value="<%= l_menuorder %>">

<table cellspacing="0" cellpadding="0" border="0" width="100%" height="100%">
	<tr>
		<td class="th2">Botones de la página</td>
		<td align="right" class="barra">
			<a class=sidebtnHLP href="Javascript:ayuda('<%= Request.ServerVariables("SCRIPT_NAME")%>');">Ayuda</a>
		</td>
	</tr>
	<tr>
		<td height="100%" colspan="2">
			<table cellspacing="0" cellpadding="0" border="0" width="100%" height="100%">
				<tr>
					<td width="50%"></td>
					<td>
						<table cellspacing="0" cellpadding="0" border="0">
							<tr>
								<td align="right"><b>P&aacute;gina:</b></td>
								<td><input type="text" name="pagina" size="35" maxlength="60" value="<%= l_pagina %>" <% If l_tipo = "M" then%>readonly <% End If%>></td>
							</tr>
							<tr>
								<td align="right"><b>Nombre:</b></td>
								<td><input type="text" name="nombre" size="35" maxlength="60" value="<%= l_nombre %>" <% If l_tipo = "M" then%>readonly <% End If%>></td>
							</tr>
							<tr>
								<td align="right" nowrap><b>Perfiles Asignados</b></td>
								<td nowrap><input type="text" name="menuaccess" size="35" maxlength="200" value="<%= l_accesos %>">
									<a class=sidebtnSHW onclick="Javascript:document.datos.menuaccess.value=''" style="cursor:hand">Borrar</a>
							</tr>
							<tr>
								<td align="right" nowrap></td>
								<td nowrap align="left">
									<%
									Set l_rs = Server.CreateObject("ADODB.RecordSet")		
									l_sql = "SELECT perfnom FROM perf_usr order by perfnom"
									rsOpen l_rs, cn, l_sql, 0 
									%>
									<select name="perfiles" style="width:235;">
									<option selected value="*">Todos</option>
									<%do until l_rs.eof%>
									<option value="<%= l_rs(0) %>"><%= l_rs(0) %></option>
									<%l_rs.MoveNext
									loop
									l_rs.Close
									set l_rs = nothing
									%>
									</select>
									<script>document.datos.perfiles.value='<%=l_primero1%>'</script>
									<a class=sidebtnSHW onclick="Javascript:agregar()" style="cursor:hand">Agregar</a>
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
		<td align="right" class="th2" colspan="2">
			<a class=sidebtnABM href="Javascript:Validar_Formulario()">Aceptar</a>
			<a class=sidebtnABM href="Javascript:window.close()">Cancelar</a>
		</td>
	</tr>
</table>
</form>
</body>
</html>

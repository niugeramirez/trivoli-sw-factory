<!--#include virtual="/turnos/shared/db/conn_db.inc"-->
<!--
-----------------------------------------------------------------------------
Archivo        : alcance_estructura_02.asp
Descripcion    : 
Modificacion   :
-----------------------------------------------------------------------------
-->
<%
Set rs = Server.CreateObject("ADODB.RecordSet")
Set rs2 = Server.CreateObject("ADODB.RecordSet")
Dim l_rs
Dim l_sql
Dim l_menuorder
Dim l_menuraiz

l_menuorder = request.QueryString("menuorder")
l_menuraiz = request.QueryString("menuraiz")
%>
<html>
<head>
<title><%= Session("Titulo")%>Ticket - Usuario: <%= l_username %></title>
<link href="/turnos/shared/css/tables3.css" rel="StyleSheet" type="text/css">
<script>
var jsSelRow = null;

function Deseleccionar(fila)
{
 fila.className = "MouseOutRow";
}
function Seleccionar(fila,cabnro)
{
 if (jsSelRow != null)
 {
  Deseleccionar(jsSelRow);
 };

 document.datos.cabnro.value = cabnro;
 fila.className = "SelectedRow";
 jsSelRow		= fila;
}
function agregar()
{
if (document.datos.menuaccess.value == "*")
	document.datos.menuaccess.value= "";
	
if ((document.datos.menuaccess.value.length + document.datos.perfiles.value.length) < 60){	
	
	if (document.datos.menuaccess.value != "")
		document.datos.menuaccess.value= document.datos.menuaccess.value + ';';
	document.datos.menuaccess.value= document.datos.menuaccess.value + document.datos.perfiles.value
}
	
}
</script>
</head>

<body bgcolor="#FFFFFF" leftmargin="0" topmargin="0" marginwidth="0">

<% 
sql = "SELECT MenuName, MenuOrder, menuaccess, menuimg FROM menumstr where menuraiz = " & l_menuraiz & " AND menuorder = " & l_menuorder
'if not username = "SUPER" then
'	sql = sql & " AND (menumstr.menuaccess LIKE '%%" & perfil & "%%' OR "
'	sql = sql & " menumstr.menuaccess = '*') "
'end if
rsOpen rs, cn, sql, 0
sql = "SELECT perfnom FROM perf_usr order by perfnom"
%>
<form name="datos" method="post">
<input type="hidden" name="menuraiz" value="<%= l_menuraiz %>">
<input type="hidden" name="menuorder" value="<%= l_menuorder %>">

<table cellpadding="0" cellspacing="0" width="100%" height="100%">
<% if not rs.eof then %>
	<tr>
		<td width="50%"></td>
		<td>
			<table cellpadding="0" cellspacing="0">
				<tr>
					<td align="right"><b>Nombre:</b></td>
					<td colspan="2"><%= rs(0) %></td>
				</tr>
				<tr>
					<td align="right"  nowrap><b>Perfiles Asignados</b></td>
					<td  nowrap><input type="text" name="menuaccess" size="40" maxlength="60" value="<%= rs(2) %>">
						&nbsp;
						<a class=sidebtnSHW onclick="Javascript:document.datos.menuaccess.value=''" style="cursor:hand">Borrar</a>
						&nbsp;
						<select name="perfiles">
						<option selected value="*">Todos</option>
						<%
						rsOpen rs2, cn, sql, 0 
						do until rs2.eof
						%>
						<option value="<%= rs2(0) %>"><%= rs2(0) %></option>
						<%
						rs2.MoveNext
						loop
						rs2.Close
						%>
						</select>
						&nbsp;
						<a class=sidebtnSHW onclick="Javascript:agregar()" style="cursor:hand">Agregar</a>
						&nbsp;
					</td>
				</tr>
				<tr>
					<td align="right" nowrap><b>Imagen Asignada</b></td>
					<td><input type="text" name="menuimg" size="40" maxlength="200"  value="<%= rs(3) %>"></td>
				</tr>
			</table>
		</td>
		<td width="50%"></td>
	</tr>
<% end if %>
<%
rs.Close
%>	
	<tr>
		<td align="right" class="th2" colspan="3">
			<a class=sidebtnABM href="Javascript:parent.modificar()">Actualizar</a>
		</td>
	</tr>
</table>

</form>
</body>
</html>

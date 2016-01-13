<!--#include virtual="/turnos/shared/db/conn_db.inc"-->
<!------------------------------------------------------------------------------
Archivo        : alcance_estructura_02.asp
Descripcion    : 
Modificacion   :
------------------------------------------------------------------------------->
<%
Set rs = Server.CreateObject("ADODB.RecordSet")
Set rs2 = Server.CreateObject("ADODB.RecordSet")
Dim l_rs
Dim l_sql
Dim l_tipo
Dim l_menuorder
Dim l_menuraiz
Dim l_nivel
Dim l_MenuName
Dim l_menuaccess
Dim l_action
Dim l_menuimg

l_menuorder = request.QueryString("menuorder")
l_menuraiz = request.QueryString("menuraiz")
l_nivel = request.QueryString("nivel")
l_tipo = request.QueryString("tipo")
%>
<html>
<head>
<title><%= Session("Titulo")%>Ticket - Usuario: <%= l_username %></title>
<link href="/turnos/shared/css/tables3.css" rel="StyleSheet" type="text/css">
<script>
var jsSelRow = null;

function Deseleccionar(fila){
	fila.className = "MouseOutRow";
}

function Seleccionar(fila,cabnro){
	if (jsSelRow != null){
		Deseleccionar(jsSelRow);
	};
	document.datos.cabnro.value = cabnro;
	fila.className = "SelectedRow";
	jsSelRow		= fila;
}

function agregar(){
	if (document.datos.menuaccess.value == "*")
		document.datos.menuaccess.value= "";
	
	if ((document.datos.menuaccess.value.length + document.datos.perfiles.value.length) < 60){	
		if (document.datos.menuaccess.value != "")
			document.datos.menuaccess.value= document.datos.menuaccess.value + ';';
		document.datos.menuaccess.value= document.datos.menuaccess.value + document.datos.perfiles.value;
	}
}
</script>
</head>

<body bgcolor="#FFFFFF" leftmargin="0" topmargin="0" marginwidth="0">

<% 
if l_tipo <> "A" then
  sql = "SELECT MenuName, MenuOrder, menuaccess, action, menuimg FROM menumstr where menuraiz = " & l_menuraiz & " AND menuorder = " & l_menuorder
  rsOpen rs, cn, sql, 0
  l_MenuName   = rs(0)
  l_menuaccess = rs(2)
  l_action     = rs(3)
  l_menuimg    = rs(4)
else
  l_MenuName   = ""
  l_MenuOrder  = ""
  l_menuaccess = ""
  l_action     = ""
  l_menuimg    = ""
end if  

sql = "SELECT perfnom FROM perf_usr order by perfnom"
%>
<form name="datos" method="post">
<input type="hidden" name="menuraiz" value="<%= l_menuraiz %>">
<input type="hidden" name="menuorder" value="<%= l_menuorder %>">
<input type="hidden" name="nivel" value="<%= l_nivel %>">
<table cellpadding="0" cellspacing="0" height="100%" width="100%" border="0">
<% if not rs.eof then %>
<tr>
	<td width="50%"></td>
	<td>
		<table cellpadding="0" cellspacing="0" border="0">
				<tr>
					<td align="right"><b>Nombre</b></td>
					<td colspan="2"><input type="text" name="nombre" size="30" maxlength="30" value="<%= l_menuname %>">
					</td>
				</tr>
				<tr>
					<td align="right" nowrap><b>Perfiles Asignados</b></td>
					<td nowrap><input type="text" name="menuaccess" size="40" maxlength="60" value="<%= l_menuaccess %>">
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
					<td align="right"><b>Acción</b></td>
					<td><input type="text" name="accion" size="76" maxlength="200"  value="<%= l_action %>"></td>
				</tr>
				<tr>
					<td align="right" nowrap><b>Imagen Asignada</b></td>
					<td><input type="text" name="menuimg" size="40" maxlength="200"  value="<%= l_menuimg %>"></td>
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
		<%
		if l_tipo = "A" then
		%>
			<a class=sidebtnABM href="Javascript:parent.altahijo()">Guar. c/ hijo</a>
			<a class=sidebtnABM href="Javascript:parent.altapar()">Guar. c/ par</a>
		<% Else  %>
			<a class=sidebtnABM href="Javascript:parent.botones()">Botones</a>
			<a class=sidebtnABM href="Javascript:parent.modificar()">Guardar</a>
		<% End If %>		
			<a class=sidebtnABM href="Javascript:document.location = 'blanc.html'">Cancelar</a>
		</td>
	</tr>
</table>
</form>
</body>
</html>

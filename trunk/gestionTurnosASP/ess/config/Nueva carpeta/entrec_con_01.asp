<% Option Explicit %>
<!--#include virtual="/serviciolocal/shared/db/conn_db.inc"-->
<% 

'Archivo: entrec_con_01.asp
'Descripción: Abm de Entregadores y recibidores
'Autor : Alvaro Bayon
'Fecha: 11/02/2005

Dim l_rs
Dim l_sql
Dim l_filtro
Dim l_orden
Dim l_sqlfiltro
Dim l_sqlorden

l_filtro = request("filtro")
l_orden  = request("orden")

if l_orden = "" then
  l_orden = " ORDER BY entdes "
end if

%>
<!DOCTYPE HTML PUBLIC "-//IETF//DTD HTML//EN">
<html>
<script src="/serviciolocal/shared/js/fn_windows.js"></script>
<script src="/serviciolocal/shared/js/fn_confirm.js"></script>
<script src="/serviciolocal/shared/js/fn_ayuda.js"></script>
<head>
<link href="/serviciolocal/shared/css/tables_gray.css" rel="StyleSheet" type="text/css">
<meta http-equiv="Content-Type" content="text/html; charset=iso-8859-1">
<title><%= Session("Titulo")%>Entregadores/Recibidores - ticket - RHPro &reg;</title>
</head>

<script>
var jsSelRow = null;

function Deseleccionar(fila){
	fila.className = "MouseOutRow";
}

function Seleccionar(fila,cabnro,des,rol){
	if (jsSelRow != null){
		Deseleccionar(jsSelRow);
	};
	document.datos.cabnro.value = cabnro;
	document.datos.descripcion.value = des;
	if (rol == 'A') {
		//	Entregador/Recibidor	
		document.datos.rol.value = 3;
	}
	else {
		if (rol == 'R') {
			// Recibidor
			document.datos.rol.value = 10;
		}
		else {
			// Entregador
			document.datos.rol.value = 9;		
		}			
	}
	
	fila.className = "SelectedRow";
	jsSelRow = fila;
}
</script>

<body leftmargin="0" rightmargin="0" topmargin="0" bottommargin="0">
<table>
    <tr>
        <th align="center">C&oacute;digo</th>
        <th>Descripci&oacute;n</th>
        <th>Tipo</th>
    </tr>
<%
Set l_rs = Server.CreateObject("ADODB.RecordSet")
'Muestro solo los ent/rec activos
l_sql = "SELECT entnro,entcod,entdes,entact,entrol"
l_sql = l_sql & " FROM tkt_entrec"
l_sql = l_sql & " WHERE entact = -1"
if l_filtro <> "" then
  l_sql = l_sql & " AND " & l_filtro & " "
end if
l_sql = l_sql & " " & l_orden
rsOpen l_rs, cn, l_sql, 0 
if l_rs.eof then%>
<tr>
	 <td colspan="3">No existen Entregadores/Recibidores</td>
</tr>
<%else
	do until l_rs.eof
	%>
	    <tr ondblclick="Javascript:parent.abrirVentana('entrec_con_02.asp?Tipo=M&cabnro=' + datos.cabnro.value,'',520,180)" onclick="Javascript:Seleccionar(this,<%= l_rs("entnro")%>,'<%= l_rs("entdes")%>','<%= l_rs("entrol") %>')">
	        <td width="10%" align="center"><%= l_rs("entcod")%></td>
	        <td width="60%" nowrap><%= l_rs("entdes")%></td>
	        <td width="30%" align="center" nowrap><% if l_rs("entrol")="E" then%>Entregador<%else if l_rs("entrol")="A" then%>Ambos<%else%>Recibidor<%end if end if%></td>
	    </tr>
	<%
		l_rs.MoveNext
	loop
end if
l_rs.Close
set l_rs = Nothing
cn.Close
set cn = Nothing
%>
</table>
<form name="datos" method="post">
<input type="Hidden" name="cabnro" value="0">
<input type="hidden" name="descripcion" value="">
<input type="hidden" name="rol" value="">
<input type="Hidden" name="orden" value="<%= l_orden %>">
<input type="Hidden" name="filtro" value="<%= l_filtro %>">
</form>
</body>
</html>

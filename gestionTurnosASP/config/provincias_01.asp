<% Option Explicit %>
<!--#include virtual="/turnos/shared/db/conn_db.inc"-->
<% 


Dim l_rs
Dim l_sql
Dim l_filtro
Dim l_orden
Dim l_sqlfiltro
Dim l_sqlorden

l_filtro = request("filtro")
l_orden  = request("orden")

if l_orden = "" then
  l_orden = " ORDER BY id"
end if

%>
<!DOCTYPE HTML PUBLIC "-//IETF//DTD HTML//EN">
<html>
<script src="/turnos/shared/js/fn_windows.js"></script>
<script src="/turnos/shared/js/fn_confirm.js"></script>
<script src="/turnos/shared/js/fn_ayuda.js"></script>
<head>
<link href="/turnos/ess/shared/css/tables_gray.css" rel="StyleSheet" type="text/css">
<meta http-equiv="Content-Type" content="text/html; charset=iso-8859-1">
<title>Obras Sociales</title>
</head>

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
	jsSelRow = fila;
}
</script>

<body leftmargin="0" rightmargin="0" topmargin="0" bottommargin="0">
<table>
    <tr>
        
        <th>Descripci&oacute;n</th>
		<!--<th>Código</th> -->
    </tr>
<%
Set l_rs = Server.CreateObject("ADODB.RecordSet")
l_sql = "SELECT * "
l_sql = l_sql & " FROM provincias "
if l_filtro <> "" then
  l_sql = l_sql & " WHERE " & l_filtro & " "
  l_sql = l_sql & " and provincias.empnro = " & Session("empnro")   
else
  l_sql = l_sql & " where provincias.empnro = " & Session("empnro")   
end if
l_sql = l_sql & " " & l_orden
rsOpen l_rs, cn, l_sql, 0 
if l_rs.eof then%>
<tr>
	 <td colspan="4">No existen Provincias</td>
</tr>
<%else
	do until l_rs.eof
	%>
	    <tr onclick="Javascript:Seleccionar(this,<%= l_rs("id")%>)">
            <td width="20%" nowrap><%= l_rs("provincia")%></td>
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
<input type="hidden" name="cabnro" value="0">
<input type="hidden" name="orden" value="<%= l_orden %>">
<input type="hidden" name="filtro" value="<%= l_filtro %>">
</form>
</body>
</html>

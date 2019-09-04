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
<title>Codigos Practicas</title>
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
        
        <th>Obra Social</th>
		<th>Practica</th>
		<th>Código</th>
    </tr>
<%
Set l_rs = Server.CreateObject("ADODB.RecordSet")
l_sql = "SELECT codigospracticas.id, obrassociales.descripcion AS osocial, practicas.descripcion AS pract, codigo "  
l_sql = l_sql & "FROM   codigospracticas, obrassociales, practicas "
l_sql = l_sql & "WHERE  codigospracticas.idpractica = practicas.id "
l_sql = l_sql & "AND    codigospracticas.idobrasocial = obrassociales.id "
l_sql = l_sql & "AND    codigospracticas.empnro = " & Session("empnro")
l_sql = l_sql & "ORDER BY 2,3"

rsOpen l_rs, cn, l_sql, 0 
if l_rs.eof then%>
<tr>
	 <td colspan="4">No existen Codigos de Practicas</td>
</tr>
<%else
	do until l_rs.eof
	%>
	    <tr ondblclick="Javascript:parent.abrirVentana('codpracticas_02.asp?Tipo=M&cabnro=' + datos.cabnro.value,'',600,300)" onclick="Javascript:Seleccionar(this,<%= l_rs("id")%>)">
            <td width="40%" nowrap align="center"><%= l_rs("osocial")%></td>
			<td width="40%" nowrap align="center"><%= l_rs("pract")%></td>
			<td width="20%" nowrap align="center"><%= l_rs("codigo")%></td>
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

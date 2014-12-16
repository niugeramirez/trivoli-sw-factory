<% Option Explicit 
response.buffer = true
%>
<!--#include virtual="/ticket/shared/db/conn_db.inc"-->
<% 
'Archivo: vendedores_con_01.asp
'Descripción: Abm de Vendedores
'Autor : Gustavo Manfrin
'Fecha: 08/02/2005
'Modificado: 

'on error goto 0

Dim l_rs
Dim l_sql
Dim l_filtro
Dim l_orden
Dim l_sqlfiltro
Dim l_sqlorden
Dim l_cont

l_filtro = request("filtro")
l_orden  = request("orden")

if l_orden = "" then
  l_orden = " ORDER BY vencordes "
end if

%>
<!DOCTYPE HTML PUBLIC "-//IETF//DTD HTML//EN">
<html>
<script src="/ticket/shared/js/fn_windows.js"></script>
<script src="/ticket/shared/js/fn_confirm.js"></script>
<script src="/ticket/shared/js/fn_ayuda.js"></script>
<head>
<link href="/ticket/shared/css/tables_gray.css" rel="StyleSheet" type="text/css">
<meta http-equiv="Content-Type" content="text/html; charset=iso-8859-1">
<title><%= Session("Titulo")%>Vendedores - Ticket</title>
</head>

<script>
var jsSelRow = null;

function Deseleccionar(fila){
	fila.className = "MouseOutRow";
}

function Seleccionar(fila,cabnro,des){
	if (jsSelRow != null){
		Deseleccionar(jsSelRow);
	};
	document.datos.cabnro.value = cabnro;
	document.datos.descripcion.value = des;
	fila.className = "SelectedRow";
	jsSelRow = fila;
}
</script>

<body leftmargin="0" rightmargin="0" topmargin="0" bottommargin="0">
<table>
    <tr>
        <th align="center">C&oacute;digo</th>
        <th>Nombre clave</th>		
        <th>Razón social</th>		
        <th>Hab.</th>				
    </tr>
<%
Set l_rs = Server.CreateObject("ADODB.RecordSet")

l_sql = "SELECT vencornro, vencorcod, vencordes, vencorrazsoc, venhab "
l_sql = l_sql & " FROM tkt_vencor "
l_sql = l_sql & " WHERE venact = -1 AND vencortip = 'V' "
if l_filtro <> "" then
  l_sql = l_sql & " AND " & l_filtro & " "
end if
l_sql = l_sql & " " & l_orden

'l_rs.maxrecords = 8000
rsOpen l_rs, cn, l_sql, 0 

if l_rs.eof then%>
<tr>
	 <td colspan="4">No existen Vendedores</td>
</tr>
<%else
	l_cont = 0
	do until l_rs.eof
		l_cont = l_cont + 1
	%>
	    <tr ondblclick="Javascript:parent.abrirVentana('vendedores_con_02.asp?Tipo=M&cabnro=' + datos.cabnro.value,'',540,260)" onclick="Javascript:Seleccionar(this,<%= l_rs("vencornro")%>,'<%= replace(l_rs("vencorrazsoc"),chr(34),"") %>')">
	        <td width="15%" align="center"><%= l_rs("vencorcod")%></td>
	        <td width="25%" nowrap><%= replace(l_rs("vencordes"),chr(34),"") %></td>
	        <td width="50%" nowrap><%= replace(l_rs("vencorrazsoc"),chr(34),"") %></td>			
            <td width="10%" align="center"><% if l_rs("venhab") = -1 then%> Si <% Else %> No <% End If%></td>			
	    </tr>
	<%
		if l_cont > 1000 then
			response.flush
			l_cont = 0
		end if
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
<input type="hidden" name="descripcion" value="">
<input type="hidden" name="orden" value="<%= l_orden %>">
<input type="hidden" name="filtro" value="<%= l_filtro %>">
</form>
</body>
</html>

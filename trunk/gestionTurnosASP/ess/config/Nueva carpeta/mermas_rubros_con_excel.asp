<% Option Explicit %>
<% Response.AddHeader "Content-Disposition", "attachment;filename=Tipos de Mermas para Rubros.xls" %>
<!--#include virtual="/ticket/shared/db/conn_db.inc"-->

<% 
'Archivo: mermas_rubros_con_excel.asp
'Descripción: ABM de Tipos de Mermas para rubros
'Autor : Alvaro Bayon
'Fecha: 15/02/2005
'Modificado: 

Dim l_rs
Dim l_sql
Dim l_filtro
Dim l_orden
Dim l_sqlfiltro
Dim l_sqlorden

l_filtro = request("filtro")
l_orden  = request("orden")

if l_orden = "" then
  l_orden = " ORDER BY lugdes "
end if

%>
<!DOCTYPE HTML PUBLIC "-//IETF//DTD HTML//EN">
<html>
<script src="/ticket/shared/js/fn_windows.js"></script>
<script src="/ticket/shared/js/fn_confirm.js"></script>
<script src="/ticket/shared/js/fn_ayuda.js"></script>
<head>

<meta http-equiv="Content-Type" content="text/html; charset=iso-8859-1">
<title><%= Session("Titulo")%>Tipos de Mermas para Rubros - Ticket</title>
</head>

<body leftmargin="0" rightmargin="0" topmargin="0" bottommargin="0">
<table>
	<tr><th colspan="3">Tipos de Mermas para Rubros</th></tr>
    <tr>
        <th>Lugar</th>
        <th>Rubro</th>
        <th>Tipo de merma</th>
    </tr>
<%
Set l_rs = Server.CreateObject("ADODB.RecordSet")
l_sql = "SELECT tipmernro,rubdes,lugcod,tipmer,forcal"
l_sql = l_sql & " FROM tkt_tipomerma "
l_sql = l_sql & " INNER JOIN tkt_rubro ON tkt_tipomerma.rubnro = tkt_rubro.rubnro"
l_sql = l_sql & " INNER JOIN tkt_lugar ON tkt_tipomerma.lugnro = tkt_lugar.lugnro"
if l_filtro <> "" then
  l_sql = l_sql & " WHERE " & l_filtro & " "
end if
l_sql = l_sql & " " & l_orden
rsOpen l_rs, cn, l_sql, 0 
if l_rs.eof then%>
<tr>
	 <td colspan="3">No existen Tipos de Mermas</td>
</tr>
<%else
	do until l_rs.eof
	%>
	    <tr ondblclick="Javascript:parent.abrirVentana('mermas_rubros_con_02.asp?Tipo=M&cabnro=' + datos.cabnro.value,'',520,200)" onclick="Javascript:Seleccionar(this,<%= l_rs("tipmernro")%>)">
	        <td width="20%" nowrap><%= l_rs("lugcod")%></td>
	        <td width="60%" nowrap><%= l_rs("rubdes")%></td>
	        <td width="20%" align="center" nowrap><%if UCase(l_rs("tipmer")) = "K" then%>Kilos<% Else %>Pesos<% End If %></td>
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

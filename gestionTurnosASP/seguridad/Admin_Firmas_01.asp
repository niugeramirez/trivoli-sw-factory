<% Option Explicit %>
<!--#include virtual="/serviciolocal/shared/db/conn_db.inc"-->
<% 
Dim l_rs
Dim l_sql

Dim l_tipoferiado
Dim l_fericompleto
dim	l_ferihoradesde 
dim	l_ferihorahasta
	
dim l_filtro
dim l_orden
l_filtro = request("filtro")
l_orden  = request("orden")

if l_orden = "" then
		l_orden = "ORDER by cysfirfecaut desc, cysfirmhora desc "
end if
%>
<!DOCTYPE HTML PUBLIC "-//IETF//DTD HTML//EN">
<html>

<head>
<link href="/serviciolocal/shared/css/tables_gray.css" rel="StyleSheet" type="text/css">
<meta http-equiv="Content-Type" http-equiv="refresh" content="text/html; charset=iso-8859-1">
<title><%= Session("Titulo")%>Feriados - Ticket</title>
</head>
<script src="/serviciolocal/shared/js/fn_windows.js"></script>
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

</script>


<body leftmargin="0" rightmargin="0" topmargin="0" bottommargin="0">
<table>
    <tr>
        <th>Enviado por </th>
        <th>Tipo </th>
        <th>Cod.</th>
        <th>Descripcion </th>
        <th>Fecha aut. </th>
        <th>Hora </th>
        <th>O.</th>
    </tr>
<%


Set l_rs = Server.CreateObject("ADODB.RecordSet")

l_sql =         "SELECT usrnombre, cysfirmas.cystipnro, cystipnombre, cysfircodext, cysfirdes, cysfirfecaut, cysfirsecuencia, cysfirmhora "
l_sql = l_sql & "FROM cysfirmas, cystipo, user_per "  
l_sql = l_sql & "WHERE cysfirmas.cysfirdestino = '" & Session("UserName") & "' and cysfirmas.cysfiryaaut = 0 and cystipo.cystipnro = cysfirmas.cystipnro and cystipo.cystipact = -1 and user_per.iduser = cysfirmas.cysfirautoriza "
		
if l_filtro <> "" then
	 l_sql = l_sql & " and " & l_filtro 
end if
	
l_sql = l_sql & l_orden	

rsOpen l_rs, cn, l_sql, 0 

	dim hdesde
Do until l_rs.eof
  if isnull(l_rs("cysfirmhora")) or l_rs("cysfirmhora") = "" then
     hdesde = "00:00"
  else
     hdesde = left(l_rs("cysfirmhora"),2) & ":" & right(l_rs("cysfirmhora"),2)
  end if
%>
    <tr onclick="javascript:Seleccionar(this)" ondblclick="Javascript:abrirVentana('Admin_Firmas_04.asp?Tipo=' + jsSelRow.cells(7).innerText + '&descripcion=' + jsSelRow.cells(3).innerText + '&codigo=' + jsSelRow.cells(2).innerText,'',545,240)">
        <td><%= l_rs("usrnombre") %></td>
        <td><%= l_rs("cystipnro") & "- " & l_rs("cystipnombre") %></td>
        <td><%= l_rs("cysfircodext") %></td>
        <td><%= l_rs("cysfirdes") %></td>
        <td><%= l_rs("cysfirfecaut") %></td>
        <td><%= hdesde %> </td>
        <td><%= l_rs("cysfirsecuencia") %></td>
        <td style='display:none'><%= l_rs("cystipnro") %></td>
    </tr>
<%    l_rs.MoveNext
	loop
l_rs.Close
cn.Close	
%>
</table>

<form name="datos" method="post">
<input type="Hidden" name="cabnro" value="0" >
<input type="Hidden" name="orden" value="<%= l_orden %>">
<input type="Hidden" name="filtro" value='<%= l_filtro %>'>
</form>

</body>
</html>

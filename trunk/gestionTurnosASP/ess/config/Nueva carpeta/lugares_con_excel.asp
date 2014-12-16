<% Option Explicit %>
<% Response.AddHeader "Content-Disposition", "attachment;filename=Lugares.xls" %>
<!--#include virtual="/serviciolocal/shared/db/conn_db.inc"-->

<% 
'Archivo: lugares_con_excel.asp
'Descripción: Consulta de Lugares
'Autor : Raul Chinestra
'Fecha: 08/02/2005
'Modificado : Raul Chinestra 02/03/2006 Se eliminaron los campos lugpro, lugbaj y se agregó el campo lugzon que indica la 
' zona comercial a la que pertenece el lugar y que se va a usar para bajar los cupos, contratos y ordenes de trabajo.


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
<script src="/serviciolocal/shared/js/fn_windows.js"></script>
<script src="/serviciolocal/shared/js/fn_confirm.js"></script>
<script src="/serviciolocal/shared/js/fn_ayuda.js"></script>
<head>
<meta http-equiv="Content-Type" content="text/html; charset=iso-8859-1">
<title><%= Session("Titulo")%>Lugares - Ticket</title>
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
		<th colspan="6">Lugares</th>
	</tr>
    <tr>
        <th>Código</th>		
        <th>Descripci&oacute;n</th>		
        <th>Localidad</th>
        <th>Provincia</th>
		<th>Zona</th>		
		<th>Estacion</th>		
		<th>Desvio</th>		
    </tr>
<%
Set l_rs = Server.CreateObject("ADODB.RecordSet")
l_sql = "SELECT lugnro, lugcod, lugdes , locdes, prodes, lugzon, estacion, desvio "
l_sql = l_sql & " FROM tkt_lugar "
l_sql = l_sql & " INNER JOIN tkt_localidad ON tkt_localidad.locnro = tkt_lugar.locnro "
l_sql = l_sql & " INNER JOIN tkt_provincia ON tkt_provincia.pronro = tkt_lugar.pronro "
if l_filtro <> "" then
  l_sql = l_sql & " WHERE " & l_filtro & " "
end if
l_sql = l_sql & " " & l_orden
rsOpen l_rs, cn, l_sql, 0 
if l_rs.eof then%>
<tr>
	 <td colspan="5">No existen Lugares</td>
</tr>
<%else
	do until l_rs.eof
	%>
	    <tr ondblclick="Javascript:parent.abrirVentana('lugares_con_02.asp?cabnro=' + datos.cabnro.value,'',520,120)" onclick="Javascript:Seleccionar(this,<%= l_rs("lugnro")%>)">
	        <td width="5%" nowrap><%= l_rs("lugcod")%></td>
	        <td width="20%" nowrap><%= l_rs("lugdes")%></td>
			<td width="20%" nowrap><%= l_rs("locdes")%></td>
			<td width="20%" nowrap><%= l_rs("prodes")%></td>
			<td width="5%" nowrap align="center"><%= l_rs("lugzon")%></td>
   		    <td width="20%" nowrap><%= l_rs("estacion")%></td>
			<td width="10%" nowrap><%= l_rs("desvio")%></td>
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

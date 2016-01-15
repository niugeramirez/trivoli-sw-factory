<%  Option Explicit %>
<!--#include virtual="/trivoliSwimming/shared/db/conn_db.inc"-->
<%

'Archivo: pol_cuenta_seg_01.asp
'Descripción: ABM de Políticas de cuenta
'Autor: Alvaro Bayon
'Fecha: 21/02/2005

 Dim l_rs
 Dim l_sql
 
 Dim l_filtro
 Dim l_orden
 
 l_filtro = request("filtro")
 l_orden  = request("orden")
 
 if l_orden = "" then
	l_orden = "ORDER BY pol_desc"
 end if
 
 Set l_rs = Server.CreateObject("ADODB.RecordSet")
%>
<!DOCTYPE HTML PUBLIC "-//IETF//DTD HTML//EN">
<html>

<head>
<link href="/trivoliSwimming/shared/css/tables_gray.css" rel="StyleSheet" type="text/css">
<meta http-equiv="Content-Type" content="text/html; charset=iso-8859-1">
<title><%= Session("Titulo")%>Pol&iacute;ticas de Cuentas - Ticket</title>
</head>

<script>
var jsSelRow = null;

function Deseleccionar(fila){
 fila.className = "MouseOutRow";
}

function Seleccionar(fila,cabnro){
 if (jsSelRow != null)
	Deseleccionar(jsSelRow);

 document.datos.cabnro.value = cabnro;
 fila.className = "SelectedRow";
 jsSelRow		= fila;
}
</script>

<body leftmargin="0" rightmargin="0" topmargin="0" bottommargin="0">
<table>
    <tr>
        <th>Descripci&oacute;n</th>
    </tr>
<%
l_sql = "SELECT pol_nro, pol_desc "
l_sql = l_sql & "FROM pol_cuenta "

if  Session("UserName") <> "sa" then
   l_sql = l_sql & " WHERE pol_desc <> 'Politica Sistemas' "
else
   l_sql = l_sql & " WHERE 1 = 1 "
end if 

if l_filtro <> "" then
'	l_sql = l_sql & "WHERE " & l_filtro & " "
	l_sql = l_sql & " AND " & l_filtro & " "
end if
l_sql = l_sql & l_orden
rsOpen l_rs, cn, l_sql, 0 
if l_rs.eof then%>
	<tr>
		 <td colspan="1">No existen Políticas de Cuentas</td>
	</tr>
<%else
	 do until l_rs.eof
	 
		%>
	    <tr ondblclick="Javascript:parent.abrirVentana('pol_cuenta_seg_02.asp?Tipo=M&pol_nro=<%=l_rs("pol_nro")%>','',720,330)" onclick="Javascript:Seleccionar(this,'<%= l_rs("pol_nro")%>')">
	        <td ><%= l_rs("pol_desc")%></td>
	    </tr>
		<%
		l_rs.MoveNext
	 loop
end if
 l_rs.Close
 cn.Close
 set l_rs = Nothing 
 set cn = Nothing 
%>
</table>
<form name="datos" method="post">
<input type="Hidden" name="cabnro" value="">
<input type="Hidden" name="orden" value="<%= l_orden %>">
<input type="Hidden" name="filtro" value="<%= l_filtro %>">
</form>
</body>
</html>

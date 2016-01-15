<% Option Explicit %>
<!--#include virtual="/trivoliSwimming/shared/db/conn_db.inc"-->
<% 

'Archivo: usuarios_seg_01.asp
'Descripción: ABM de usuarios 
'Autor: Alvaro Bayon
'Fecha: 21/02/2005

Dim l_rs
Dim l_sql
Dim l_rs1
Dim l_sql1

Dim l_filtro
Dim l_orden
Dim l_sqlfiltro
Dim l_sqlorden

Dim l_iduser
Dim l_perfnro
Dim l_empleado
Dim l_libreria
Dim l_MRUOrden
Dim l_MRUCant
Dim l_usrtipsg

Dim l_nombre ' variable para el browse 

l_filtro = request("filtro")
l_orden  = request("orden")

if l_orden = "" then
  l_orden = " ORDER BY iduser"
end if

%>
<!DOCTYPE HTML PUBLIC "-//IETF//DTD HTML//EN">
<html>

<head>
<link href="/trivoliSwimming/shared/css/tables_gray.css" rel="StyleSheet" type="text/css">
<meta http-equiv="Content-Type" content="text/html; charset=iso-8859-1">
<title><%= Session("Titulo")%>Usuarios - Ticket</title>
</head>

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
        <th>Id. Usuario</th>
        <th>Nombre Usuario</th>
        <th>Perfil</th>
		<th>Email</th>
    </tr>
<%

if l_orden = "" then
	l_orden = "ORDER BY iduser"
end if

Set l_rs = Server.CreateObject("ADODB.RecordSet")

l_sql = "SELECT iduser,usrnombre, MRUOrden, MRUCant, perfnom, usremail "
l_sql = l_sql & " FROM user_per"
l_sql = l_sql & " INNER JOIN perf_usr ON user_per.perfnro = perf_usr.perfnro "

'response.write "username " &  Session("UserName")

if  Session("UserName") <> "sa" then
  l_sql = l_sql & " WHERE iduser <> 'sa' "
else
  l_sql = l_sql & " WHERE 1 = 1 "
end if 

'response.write "FILTRO " & l_filtro
if l_filtro <> "" then
'  l_sql = l_sql & "WHERE " & l_filtro & " "
  l_sql = l_sql & " AND " & l_filtro & " "
end if
l_sql = l_sql & l_orden
rsOpen l_rs, cn, l_sql, 0 

do until l_rs.eof
%>
    <tr ondblclick="Javascript:parent.abrirVentanaVerif('usuarios_seg_02.asp?Tipo=M&iduser=<%=l_rs("iduser")%>','',550,350)" onclick="Javascript:Seleccionar(this,'<%= l_rs("iduser")%>')">
        <td width="25%"><%= l_rs("iduser")%></td>
        <td width="55%"><%= l_rs("usrnombre")%></td>
        <td width="20%"><%= l_rs("perfnom")%></td>
		<td width="20%"><%= l_rs("usremail")%></td>
    </tr>
<% l_rs.MoveNext
loop

l_rs.Close
cn.Close
set l_rs = Nothing 
set cn = Nothing 
%>
</table>
<form name="datos" method="post">
<input type="Hidden" name="cabnro" value="<%= l_iduser %>" >
<input type="Hidden" name="orden" value="<%= l_orden %>">
<input type="Hidden" name="filtro" value="<%= l_filtro %>">
</form>
</body>
</html>

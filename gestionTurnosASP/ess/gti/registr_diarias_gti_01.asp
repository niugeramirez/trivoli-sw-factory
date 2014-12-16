<% Option Explicit %>
<!--#include virtual="/serviciolocal/ess/shared/inc/const.inc"-->
<!--#include virtual="/serviciolocal/shared/db/conn_db.inc"-->
<!--#include virtual="/serviciolocal/shared/inc/fecha.inc"-->
<!--#include virtual="/serviciolocal/ess/shared/inc/accesoESS.inc"-->
<!--
-----------------------------------------------------------------------------
Archivo        : registr_diarias_gti_01.asp
Descripcion    : Modulo que se encarga de mostrar es listado HTML de registraciones
Modificacion   :
   10/10/2003 - Scarpa D. - Si la registracion es nula la muestra como desconocida
-----------------------------------------------------------------------------
-->
<% 
on error goto 0

Dim l_rs
Dim l_sql
Dim l_reghora
Dim l_regfecha

Dim l_filtro
Dim l_orden
Dim l_sqlfiltro
Dim l_sqlorden

Dim l_fechadesde
Dim l_fechahasta
Dim l_ternro

l_filtro = request("filtro")
l_orden  = request("orden")

l_ternro = l_ess_ternro
l_fechadesde = request.querystring("fechadesde")
l_fechahasta  = request.querystring("fechahasta")

if l_orden = "" then
   l_orden = 2  'orden por defecto Fecha, hora
end if

%>
<!DOCTYPE HTML PUBLIC "-//IETF//DTD HTML//EN">
<html>

<head>
<link href="../<%=c_estiloTabla %>" rel="StyleSheet" type="text/css">
<meta http-equiv="Content-Type" content="text/html; charset=iso-8859-1">
<title>Registraciones - Gesti&oacute;n de Tiempos - RHPro &reg;</title>
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
    <tr nowrap>
        <th align="center">Fecha</th>
        <th align="center">D&iacute;a</th>
        <th align="center">Hora</th>
        <th align="left">Reloj</th>
        <th align="center">E/S</th>
        <th align="center">Estado</th>
        <th align="center">Manual</th>
    </tr>
<%

select case l_orden
  case "-1"  l_sqlorden = "ORDER BY reldabr DESC, regfecha "
  case "-2"  l_sqlorden = "ORDER BY regfecha DESC, reghora"
  case "1"  l_sqlorden = "ORDER BY reldabr, regfecha "
  case "2"  l_sqlorden = "ORDER BY regfecha, reghora"
  case else l_sqlorden = ""
end select

Set l_rs = Server.CreateObject("ADODB.RecordSet")
l_sql = "SELECT regnro, regfecha, reghora, gti_reloj.reldabr, regestado, regentsal, regmanual "
l_sql = l_sql & "FROM gti_registracion  LEFT JOIN gti_reloj ON gti_registracion.relnro=gti_reloj.relnro "
l_sql = l_sql & "WHERE ternro="& l_ternro & " and regfecha>=" & cambiafecha(l_fechadesde,"YMD",true) & " and regfecha<="& cambiafecha(l_fechahasta,"YMD",true) &" "
if l_filtro <> "" then
  l_sql = l_sql & " and " & l_filtro & " "
end if
l_sql = l_sql & l_sqlorden

rsOpen l_rs, cn, l_sql, 0

do until l_rs.eof
	l_reghora = l_rs("reghora")
	l_regfecha= l_rs("regfecha")
%>
    <tr onclick="Javascript:Seleccionar(this,<%= l_rs("regnro")%>)">
        <td nowrap align="center"><%= l_regfecha %></td>
		<td nowrap align="center"><%= calculardia(l_regfecha)%></td>
        <td  nowrap align="center"><%= mid(l_reghora,1,2)&":"& mid(l_reghora,3,2)%></td>
        <td nowrap align="left"><%= l_rs("reldabr") %></td>    
		<%if isNull(l_rs("regentsal")) then%>
          <td nowrap align="center">D</td>		
		<%else%>
          <td nowrap align="center"><%= l_rs("regentsal")%></td>
		<%end if%>
        <td nowrap align="center"><%= l_rs("regestado")%></td>
        <td align="center"><% if l_rs("regmanual")= -1 then
									response.write "Si"
								else
									response.write "No"
								end if %></td>    
	</tr>
<%
	l_rs.MoveNext
loop
l_rs.Close

set l_rs = Nothing
cn.Close
set cn = Nothing
%>
</table>

<% if l_fechadesde<>"" then %>
	<script> parent.datos.fechadesde.value="<%= l_fechadesde %>";</script>
<% end if %>
<% if l_fechahasta<>"" then %>
	<script> parent.datos.fechahasta.value="<%= l_fechahasta %>";</script>
<% end if %>


<form name="datos" method="post">
<input type="Hidden" name="cabnro" value="0">
<input type="Hidden" name="orden" value="<%= l_orden %>">
<input type="Hidden" name="filtro" value="<%= l_filtro %>">
</form>
</body>
</html>

<% Option Explicit %>
<%
Response.AddHeader "Content-Disposition", "attachment;filename=Configuración email.xls" 
Response.ContentType = "application/vnd.ms-excel" %>
<!--#include virtual="/serviciolocal/shared/db/conn_db.inc"-->
<%
'Archivo        : conf_mil_00.asp
'Descripcion    : Modulo que se encarga de admin. los servidores de mail
'Creador        : Lisandro Moro
'Fecha Creacion : 08/03/2005
'Modificacion   :

Dim l_rs
Dim l_sql

Dim l_tiptabnro
Dim l_tiptabdesc

Dim l_filtro
Dim l_orden

l_filtro = request("filtro")
l_orden  = request("orden")

if l_orden = "" then
  l_orden = " ORDER BY cfgemailnro ASC"  'orden por número asc
end if

%>
<!DOCTYPE HTML PUBLIC "-//IETF//DTD HTML//EN">
<html>

<head>
<meta http-equiv="Content-Type" http-equiv="refresh" content="text/html; charset=iso-8859-1">
<title><%= Session("Titulo")%>Configuraci&oacute;n se Servicios de Mail - RHPro &reg;</title>
</head>

<body leftmargin="0" rightmargin="0" topmargin="0" bottommargin="0">
<table>
    <tr>
        <th>C&oacute;digo</th>
        <th>Descripci&oacute;n</th>
        <th>Origen</th>		
        <th>Host</th>				
        <th>Puerto</th>						
        <th>Estado</th>		
    </tr>
<%

	

Set l_rs = Server.CreateObject("ADODB.RecordSet")
l_sql = "SELECT * "
l_sql = l_sql & " FROM  conf_email "
l_sql = l_sql & " WHERE 1=1 "

if l_filtro <> "" then
  l_sql = l_sql & " AND " & l_filtro & " "
end if

l_sql = l_sql & " " & l_orden

rsOpen l_rs, cn, l_sql, 0 
if l_rs.eof then%>
<tr>
	 <td colspan="6">No hay datos</td>
</tr>
<%else
	do until l_rs.eof
	%>
	<tr>
		<td width="10%"><%=l_rs("cfgemailnro")%></td>
		<td width="10%"><%=l_rs("cfgemaildesc")%></td>
		<td width="30%"><%=l_rs("cfgemailfrom")%> </td>
		<td width="30%"><%=l_rs("cfgemailhost")%> </td>		
		<td width="30%"><%=l_rs("cfgemailport")%> </td>				
		<%if CInt(l_rs("cfgemailest")) = -1 then %>
		<td width="30%">Activa</td>				
		<%else%>
		<td width="30%">Inactiva</td>						
		<%end if%>
	</tr>
	<%l_rs.MoveNext
	loop
end if 
l_rs.Close
cn.Close	
%>
</table>

</body>
</html>

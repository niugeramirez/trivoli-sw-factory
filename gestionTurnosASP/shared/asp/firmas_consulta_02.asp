<% Option Explicit %>
<!--#include virtual="/serviciolocal/shared/db/conn_db.inc"-->
<% Response.AddHeader "Content-Disposition", "attachment;filename=consulta_firmas.xls" %> 
<%
'Archivo	: firmas_consulta_01.asp
'Descripción: Consulta de Firmas Excel
'Autor		: CCRossi
'Fecha		: 03-02-2004
'Modificacion: 
'------------------------------------------------------------------------------------

'Variables base de datos
 Dim l_rs
 Dim l_sql
 
 'Variables filtro y orden
Dim l_filtro
Dim l_orden
 
'Tomar parametros
l_filtro = request("filtro")
l_orden  = request("orden")

Dim l_cysfircodext
l_cysfircodext = Request.QueryString("cysfircodext")
Dim l_cystipnro
l_cystipnro = Request.QueryString("cystipnro")

if l_orden = "" then
   l_orden = " ORDER BY cysfirmas.cysfirfecaut DESC"
end if

'Body 
%>
<!DOCTYPE HTML PUBLIC "-//IETF//DTD HTML//EN">
<html>
<head>
<meta http-equiv="Content-Type" http-equiv="refresh" content="text/html; charset=iso-8859-1">
<title><%= Session("Titulo")%>Consulta de Firmas - RHPro &reg;</title>
</head>
<script>
</script>
<body leftmargin="0" rightmargin="0" topmargin="0" bottommargin="0">
<table>
    <tr>
        <th nowrap>Fecha</th>
        <th nowrap>Hora</th>
        <th nowrap>Autorizado Por</th>				
        <th nowrap>Tipo</th>				
        <th nowrap>Fin de Firma</th>				
    </tr>
<%
Set l_rs = Server.CreateObject("ADODB.RecordSet")
l_sql = "SELECT cysfirmas.cysfirfecaut, cysfirmas.cysfirmhora , cysfirmas.cysfirautoriza, "
l_sql = l_sql & " cystipo.cystipnombre, cysfirmas.cysfirfin "
l_sql = l_sql & " FROM  cysfirmas "
l_sql = l_sql & " INNER JOIN  cystipo ON cysfirmas.cystipnro = cystipo.cystipnro "
l_sql = l_sql & " WHERE  cysfirmas.cysfircodext = " & l_cysfircodext
l_sql = l_sql & " AND    cysfirmas.cystipnro = " & l_cystipnro
if l_filtro <> "" then
	 l_sql = l_sql & " AND " & l_filtro 
end if

l_sql = l_sql & " " & l_orden

'response.write l_sql
rsOpen l_rs, cn, l_sql, 0
if l_rs.eof then%>
<tr>
	 <td colspan="5">No hay Firmas para el registro seleccionado.</td>
</tr>
<%else
	do until l_rs.eof%>
	
<tr>
	<td nowrap align=center><%=l_rs("cysfirfecaut")%> </td>
	<td nowrap align=center><%=l_rs("cysfirmhora")%></td>
	<td nowrap align=center><%=l_rs("cysfirautoriza")%></td>
	<td nowrap><%=l_rs("cystipnombre")%></td>
	<%if CInt(l_rs("cysfirfin")) = -1 then
	    response.write "	<td align=center nowrap>SI</td>"
	  else 
	    response.write "	<td align=center nowrap>NO</td>"	
	  end if
	%>	
</tr>
	<%l_rs.MoveNext
	loop
end if ' del if l_rs.eof 

l_rs.Close
set l_rs = nothing

cn.Close	
set cn = nothing
%>
</table>
</body>
</html>
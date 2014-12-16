<% Option Explicit %>
<!--#include virtual="/serviciolocal/shared/db/conn_db.inc"-->
<!--#include virtual="/serviciolocal/shared/inc/fecha.inc"-->
<!--#include virtual="/serviciolocal/ess/shared/inc/accesoESS.inc"-->
<%Response.AddHeader "Content-Disposition", "attachment;filename=nota.xls" %>
<% 
'---------------------------------------------------------------------------------
'Archivo	: nota_adp_excel.asp
'Descripción: salida excel del nota
'Autor		: Claudia Cecilia Rossi
'Fecha		: 30-08-2003
'Modificado	: 08-11-2005 - Leticia A. - Adecuarlo para Autogestion.
'			: 08-11-2005 - Leticia A. - usar FechaISO para mostrar las fechas 
'Modificado  : 09/09/2006 Raul Chinestra - se agregó el campo Revisada a las Notas
'----------------------------------------------------------------------------------

'Variables base de datos
 Dim l_rs
 Dim l_sql

'uso local

'Variables filtro y orden
 dim l_filtro
 dim l_filtro2
 dim l_orden
 
'var parametro de entrada 
 dim l_ternro
  
'Tomar parametros
 l_ternro = l_ess_ternro
 l_filtro = request("filtro")
 l_orden  = request("orden")

'Body 
 if l_orden = "" then
	l_orden = " ORDER BY tnodesabr"
 end if
 
%>
<html>

<head>
<meta http-equiv="Content-Type" http-equiv="refresh" content="text/html; charset=iso-8859-1">
<title>Notas - Administraci&oacute;n de Personal - RHPro &reg;</title>
</head>
<body leftmargin="0" rightmargin="0" topmargin="0" bottommargin="0">
<table>
   <tr>
        <th colspan=5>Notas</th>
    </tr>
    <tr>
        <th>Tipo Nota</th>
		<th>Remitente</th>
		<th>Fecha Ingr.</th>
		<th>Revisada</th>		
		<th>Fecha Vto.</th>
        <th>Nota Abr.</th>
    </tr>
<%
Set l_rs = Server.CreateObject("ADODB.RecordSet")
l_sql = "SELECT notanro, notas_ter.tnonro, tnodesabr, notfecalta, notfecvenc, notremitente, "  
l_sql = l_sql & " notatxt  "
l_sql = l_sql & " FROM notas_ter "
l_sql = l_sql & " INNER JOIN tiponota ON tiponota.tnonro = notas_ter.tnonro "
l_sql = l_sql & " WHERE notas_ter.ternro =  " & l_ternro
l_sql = l_sql & " AND   tiponota.tnoconfidencial = 0 "
if l_filtro <> "" then
 l_sql = l_sql & " AND " & l_filtro
end if
l_sql = l_sql & " " & l_orden	
rsOpen l_rs, cn, l_sql, 0 

if l_rs.eof then%>
<tr>
	 <td colspan="5">No hay datos.</td>
</tr>
<%else%>
	<%	do until l_rs.eof %>
	<tr onclick="Javascript:Seleccionar(this,<%=l_rs("notanro")%>)">
		<td width="20%"><%=l_rs("tnodesabr")%></td>
		<td width="20%" align="center"><%=l_rs("notremitente")%></td>
		<td width="15%" align="center"><%=fechaISO(l_rs("notfecalta"))%></td>
		<td width="5%" align="center" ><% if l_rs("notrev") = -1 then response.write "SI" else response.write "NO" end if %></td>						
		<td width="15%"><%=fechaISO(l_rs("notfecvenc"))%></td>
		<td width="30%"><%=left(l_rs("notatxt"),50)%> </td>
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

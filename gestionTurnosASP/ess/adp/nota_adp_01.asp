<% Option Explicit %>
<!--#include virtual="/ess/ess/shared/inc/const.inc"-->
<!--#include virtual="/serviciolocal/shared/db/conn_db.inc"-->
<!--#include virtual="/serviciolocal/ess/shared/inc/accesoESS.inc"-->
<% on error goto 0
'---------------------------------------------------------------------------------
'Archivo	: tiponota_adp_01.asp
'Descripción: browse de datos de tiponota
'Autor		: Claudia Cecilia Rossi
'Fecha		: 30-08-2003
'Modificado	: 08-11-05 - Leticia A. - Adecuarlo para Autogestion.
'Modificado	: 11-09-2006-RCH - Se agrego el campo revisada
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
 l_filtro = request("filtro")
 l_orden  = request("orden")

l_ternro = l_ess_ternro

 
'Body 
 if len(l_filtro) <> 0 then
	if left(l_filtro,1) <> "'" then
		l_filtro2 = "'" & l_filtro & "'"
	else
		l_filtro2 =  mid(l_filtro,2,len(request("filtro")) - 1)
	end if	
 end if	
 if l_orden = "" then
	l_orden = " ORDER BY tnodesabr"
 end if
%>
<!DOCTYPE HTML PUBLIC "-//IETF//DTD HTML//EN">
<html>

<head>
<link href="../<%=c_estiloTabla %>" rel="StyleSheet" type="text/css">
<meta http-equiv="Content-Type" http-equiv="refresh" content="text/html; charset=iso-8859-1">
<title>Notas - Administraci&oacute;n de Personal - RHPro &reg;</title>
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
l_sql = l_sql & " notatxt, notrev  "
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
	 <td colspan="5">No hay datos</td>
</tr>
<%else%>
	<%
	do until l_rs.eof%>
	
	<tr ondblclick="Javascript:parent.abrirVentanaVerif('nota_adp_02.asp?Tipo=M&ternro=' + document.datos.ternro.value+'&notanro='+document.datos.cabnro.value,'',600,360)" onclick="Javascript:Seleccionar(this,<%=l_rs("notanro")%>)">
		<td width="20%"><%=l_rs("tnodesabr")%></td>
		<td width="20%" align="center"><%=l_rs("notremitente")%></td>
		<td width="15%" align="center"><%=l_rs("notfecalta")%></td>
		<td width="5%" align="center" ><% if l_rs("notrev") = -1 then response.write "SI" else response.write "NO" end if %></td>				
		<td width="15%"><%=l_rs("notfecvenc")%></td>
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

<form name="datos" method="post">
<input type="Hidden" name="cabnro" value="0" >
<input type="Hidden" name="ternro" value="<%=l_ternro%>" >
<input type="Hidden" name="orden" value="<%= l_orden %>">
<input type="Hidden" name="filtro" value="<%= l_filtro %>">
</form>

</body>
</html>

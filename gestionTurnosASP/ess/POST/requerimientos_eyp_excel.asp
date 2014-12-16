<% Option Explicit %>
<% Response.AddHeader "Content-Disposition", "attachment;filename=Requerimientos.xls" %>
<!--#include virtual="/ess/ess/shared/inc/const.inc"-->
<!--#include virtual="/serviciolocal/shared/db/conn_db.inc"-->
<!--#include virtual="/serviciolocal/ess/shared/inc/accesoESS.inc"-->
<% on error goto 0
'---------------------------------------------------------------------------------
'Archivo	: requerimientos_eyp_01.asp
'Descripción: browse de datos de requerimientos
'Autor		: Raul Chinestra
'Fecha		: 12/09/2006
' Modificado  : 12/09/2006 Raul Chinestra - se agregó Requerimientos de Personal en Autogestión   
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
 dim l_empleg
 
'Tomar parametros
 l_filtro = request("filtro")
 l_orden  = request("orden")

l_ternro = l_ess_ternro
l_empleg = l_ess_empleg

 
'Body 
 if len(l_filtro) <> 0 then
	if left(l_filtro,1) <> "'" then
		l_filtro2 = "'" & l_filtro & "'"
	else
		l_filtro2 =  mid(l_filtro,2,len(request("filtro")) - 1)
	end if	
 end if	
 if l_orden = "" then
	l_orden = " ORDER BY reqpernro"
 end if
%>
<!DOCTYPE HTML PUBLIC "-//IETF//DTD HTML//EN">
<html>

<head>
<meta http-equiv="Content-Type" http-equiv="refresh" content="text/html; charset=iso-8859-1">
<title>Requerimientos - Empleos y Postulantes - RHPro &reg;</title>
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
        <th>Código</th>
		<th>Descripción</th>
		<th>Fecha Solicitud</th>
    </tr>
<%
Set l_rs = Server.CreateObject("ADODB.RecordSet")
l_sql = "SELECT reqpernro, reqperdesabr, reqpersolfec "  
l_sql = l_sql & " FROM pos_reqpersonal "
l_sql = l_sql & " WHERE reqpersolpor =  " & l_empleg
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
	
	<tr ondblclick="Javascript:parent.abrirVentanaVerif('requerimientos_eyp_02.asp?Tipo=M&ternro=' + document.datos.ternro.value+'&notanro='+document.datos.cabnro.value,'',600,360)" onclick="Javascript:Seleccionar(this,<%=l_rs("reqpernro")%>)">
		<td width="10%" align="center"><%=l_rs("reqpernro")%></td>
		<td width="70%" align="left"><%=l_rs("reqperdesabr")%></td>
		<td width="15%" align="center"><%=l_rs("reqpersolfec")%></td>
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
<input type="hidden" name="cabnro" value="0" >
<input type="Hidden" name="ternro" value="<%=l_ternro%>" >
<input type="Hidden" name="orden" value="<%= l_orden %>">
<input type="Hidden" name="filtro" value="<%= l_filtro %>">
</form>

</body>
</html>

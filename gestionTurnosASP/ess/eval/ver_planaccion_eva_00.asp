<% Option Explicit %>
<!--#include virtual="/ess/ess/shared/inc/const.inc"-->
<!--#include virtual="/serviciolocal/shared/db/conn_db.inc"-->
<!--#include virtual="/serviciolocal/shared/inc/const_eva.inc"-->
<% 
'--------------------------------------------------------------------------
'Archivo       : ver_planaccion_eva_00.asp
'Descripcion   : ver planes
'Creacion      : 27-05-2004
'Autor         : CCRossi
'Modificacion  : 08-02-2005 * adecuacionpara Codelco
'Modificacion  : 21-03-2005 * cambiar letra en clase .rev
'                13-10-2005 - Leticia Amadio -  Adecuacion a Autogestion
'		         21-08-2007 - Diego Rosso - Se agrego src="blanc.asp" para https
'--------------------------------------------------------------------------

Dim l_rs
Dim l_sql
Dim l_filtro
Dim l_orden

Dim l_evldrnro
Dim col
col = 2
if ccodelco<>-1 then
	col= col + 1
end if
	
l_evldrnro = request.querystring("evldrnro")

if l_orden = "" then
  l_orden = " ORDER BY evaplnro "
end if

%>
<!DOCTYPE HTML PUBLIC "-//IETF//DTD HTML//EN">
<html>

<head>
<link href="../<%=c_estiloTabla  %>" rel="StyleSheet" type="text/css">
<meta http-equiv="Content-Type" content="text/html; charset=iso-8859-1">
<title>Gesti&oacute;n de Desempe&ntilde;o - RHPro &reg;</title>
</head>
<script src="/serviciolocal/shared/js/fn_fechas.js"></script>
<script>
</script>
<style>
.rev
{
	font-size: 12;
	border-style: none;
}
</style>
<body leftmargin="0" rightmargin="0" topmargin="0" bottommargin="0">
<table>
    <tr style="border-color :CadetBlue;">
    <%if ccodelco=-1 then%>
		<th class="th2">Seguimiento</th>
	<%else%>	
        <th class="th2">Aspecto a Mejorar</th>
        <th class="th2">Plan de Acci&oacute;n</th>
     <%end if%> 
        <th class="th2">Fecha <%if ccodelco<>-1 then%>de Revisión<%end if%></th>
    </tr>
<form name="datos" method="post">
<%

Set l_rs = Server.CreateObject("ADODB.RecordSet")
l_sql = "SELECT evaplnro, aspectomejorar, planaccion, planfecharev "
l_sql = l_sql & "FROM evaplan "
l_sql = l_sql & "WHERE evldrnro =" & l_evldrnro
rsOpen l_rs, cn, l_sql, 0 
if l_rs.eof then
%>
    <td colspan="<%=col%>">No hay datos cargados.</td>
<%
end if
do until l_rs.eof
%>
    <tr>
        <td align=center valign=middle>
			<textarea class="rev" readonly style="background : #e0e0de;" name="aspectomejorar<%=l_rs("evaplnro")%>"  maxlength=200 size=200 cols=40 rows=4><%=trim(l_rs("aspectomejorar"))%></textarea>
		</td>
		<%if ccodelco<>-1 then%>
        <td align=center valign=middle>
			<textarea class="rev" readonly style="background : #e0e0de;" name="planaccion<%=l_rs("evaplnro")%>"  maxlength=200 size=200 cols=40 rows=4><%=trim(l_rs("planaccion"))%></textarea>
		</td>
		<%else%>
			<input type=hidden name="planaccion<%=l_rs("evaplnro")%>">	
		<%end if%>
        <td align=center valign=middle>
			<input class="rev" readonly style="background : #e0e0de;" type="text" name="planfecharev<%=l_rs("evaplnro")%>" size="10" maxlength="10" value="<%=l_rs("planfecharev")%>">
		</td>
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
<iframe name="grabar" src="blanc.asp" style="visibility:hidden;width:0;height:0">
</iframe>

<input type="Hidden" name="cabnro" value="0">
</form>
</body>
</html>

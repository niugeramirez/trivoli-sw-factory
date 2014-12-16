<% Option Explicit %>
<!--#include virtual="/ess/ess/shared/inc/const.inc"-->
<!--#include virtual="/serviciolocal/shared/db/conn_db.inc"-->
<!--#include virtual="/serviciolocal/shared/inc/const_eva.inc"-->
<% 
'=====================================================================================
'Archivo  : ver_evalobjetivos_eva_00.asp
'Objetivo : Ver Evaluacion de objetivos de evaluacion
'Fecha	  : 02-06-2004
'Autor	  : CCRossi
'            13-10-2005 - Leticia Amadio -  Adecuacion a Autogestion
'=====================================================================================
 Dim l_rs
 Dim l_rs1
 Dim l_sql
 Dim l_filtro
 Dim l_orden

'parametros
 Dim l_evldrnro
 Dim l_evapernro 'periodo de evaluacion
 
 l_evldrnro = request.querystring("evldrnro")
 l_evapernro = request.querystring("evapernro")

 if l_orden = "" then
  l_orden = " ORDER BY evaobjnro "
 end if

%>
<!DOCTYPE HTML PUBLIC "-//IETF//DTD HTML//EN">
<html>

<head>
<link href="../<%=c_estiloTabla  %>" rel="StyleSheet" type="text/css">
<meta http-equiv="Content-Type" content="text/html; charset=iso-8859-1">
<title>Gesti&oacute;n de Desempe&ntilde;o - RHPro &reg;</title>
</head>

<script>

</script>

<body leftmargin="0" rightmargin="0" topmargin="0" bottommargin="0">
<table>
    <tr>
        <th align=center class="th2">Descripci&oacute;n</th>
        <%if cformed=-1 then%>
        <th align=center class="th2">Forma de Medici&oacute;n</th>
        <%else%>
        <th align=center class="th2">&nbsp;</th>
        <%end if%>
        <th class="th2">&nbsp;</th>        
    </tr>
<form name="datos" method="post">
<%

Set l_rs = Server.CreateObject("ADODB.RecordSet")
l_sql = "SELECT evaobjetivo.evaobjnro,evaperfijo, evapernroeva, evaobjdext,evaobjformed, evldrnro, evatrnro "
l_sql = l_sql & "FROM evaobjetivo "
l_sql = l_sql & " INNER JOIN evaluaobj ON evaluaobj.evaobjnro = evaobjetivo.evaobjnro"
l_sql = l_sql & " WHERE evaluaobj.evldrnro =" & l_evldrnro
rsOpen l_rs, cn, l_sql, 0 
'Response.Write l_sql
if l_rs.EOF then
%>
    <tr>
        <td align=center colspan=4><b>No hay se han definido Objetivos.</b></td>
    </tr>
<%
else
do until l_rs.eof
%>
    <tr>
        <td align=center>
			<textarea disabled readonly name="evaobjdext<%=l_rs("evaobjnro")%>"  maxlength=200 size=200 cols=30 rows=4><%=trim(l_rs("evaobjdext"))%></textarea>
		</td>
        <td align=center>
			<%if cformed=-1 then%>
        	<textarea name="evaobjformed<%=l_rs("evaobjnro")%>"  maxlength=200 size=200 cols=30 rows=4><%=trim(l_rs("evaobjformed"))%></textarea>
			<%else%>
			<input name="evaobjformed<%=l_rs("evaobjnro")%>" type=hidden value="<%=trim(l_rs("evaobjformed"))%>">
			<%end if%>
		</td>
        <td nowrap>
			<%'BUSCAR la descripcion de evaresu  ----------------------------
		    Set l_rs1 = Server.CreateObject("ADODB.RecordSet")
			l_sql = "SELECT  evatipresu.evatrnro, evatipresu.evatrvalor, evatipresu.evatrdesabr "
			l_sql = l_sql & " FROM evatipresu  "
			l_sql = l_sql & " WHERE evatrtipo=2 "
			l_sql = l_sql & " order by evatrvalor "
			rsOpen l_rs1, cn, l_sql, 0%>
			<select disabled readonly name="evatrnro<%=l_rs("evaobjnro")%>">
			<%do while not l_rs1.eof%>
				<option value=<%=l_rs1("evatrnro")%>><%=l_rs1("evatrvalor")%>&nbsp;-&nbsp;<%=l_rs1("evatrdesabr")%></option>
			<%l_rs1.MoveNext
			loop 
			l_rs1.Close
			set l_rs1 = nothing%>
			</select>
			<script>document.datos.evatrnro<%=l_rs("evaobjnro")%>.value='<%=l_rs("evatrnro")%>'</script>
		</td>
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

<input type="Hidden" name="cabnro" value="0">
</form>
</body>
</html>

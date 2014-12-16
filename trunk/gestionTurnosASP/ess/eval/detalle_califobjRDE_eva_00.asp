<%Option Explicit %>
<!--#include virtual="/ess/ess/shared/inc/const.inc"-->
<!--#include virtual="/serviciolocal/shared/inc/sec.inc"-->
<!--#include virtual="/serviciolocal/shared/inc/const.inc"-->
<!--#include virtual="/serviciolocal/shared/db/conn_db.inc"-->
<!--#include virtual="/serviciolocal/shared/inc/const_eva.inc"-->
<%
'================================================================================
'Archivo		: detalle_califobjRDE_eva_ag_00.asp
'Descripción	: Muestra el detalle de las evaluaciones de los objs  RDE de un dado proyecto y de un empleado
'Autor			:01  -06-2005
'Fecha			: L Amadio
'Modificado		: 03-08-2005 - L.A. - Cambiar cod de proyecto por cod de evento.
'            Leticia Amadio - 13-10-2005 - Adecuacion a Autogestion	
'================================================================================
on error goto 0

' Variables
  dim l_empleado
  dim i

' de base de datos  
  Dim l_sql
  Dim l_rs
  Dim l_rs2
  Dim l_rs1
  
  dim l_evaproynro
  dim l_ternro  
  dim l_evento

' parametros de entrada-----------------------------
  l_evaproynro =  Request.QueryString("evaproynro")
  l_ternro = Request.QueryString("ternro")
%>

<html>
<head>
<link href="../<%=c_estilo %>" rel="StyleSheet" type="text/css">
<meta http-equiv="Content-Type" content="text/html; charset=iso-8859-1">
<title>Gesti&oacute;n de Desempe&ntilde;o - RHPro &reg;</title>
<script src="/serviciolocal/shared/js/fn_windows.js"></script>
<script src="/serviciolocal/shared/js/fn_confirm.js"></script>
<script src="/serviciolocal/shared/js/fn_ayuda.js"></script>
<style>
.rev {
	font-size: 10;
	border-style: none;
}
</style>
</head>

<script>
</script>
<body leftmargin="0" topmargin="0" rightmargin="0">
<form name="datos">
<% 	
Set l_rs1 = Server.CreateObject("ADODB.RecordSet")
l_sql = " SELECT evaevenro FROM evaevento WHERE evaproynro="& l_evaproynro
rsOpen l_rs1, cn, l_sql, 0 
l_evento = l_rs1("evaevenro")
l_rs1.Close

l_sql= " SELECT terape, terape2,ternom, ternom2 "
l_sql = l_sql & " FROM tercero  WHERE ternro= " & l_ternro
rsOpen l_rs1, cn, l_sql, 0 
l_empleado = l_rs1("terape") & " " &  l_rs1("terape2") & " " & l_rs1("ternom") &  " "  & l_rs1("ternom2")
l_rs1.Close
%>

<table border="0" cellpadding="0" cellspacing="1" width="100%">
<tr height="20">
	<td colspan="4" align="center"><b>Calificaci&oacute;n de Objetivos RDE </b></td> 
</tr>
<tr height="20">
	<td colspan="4"><b>Empleado</b>: <%=l_empleado%></td> 
</tr>
<tr height="20">
	<td colspan="4"><b> Evento nro:</b> <%=l_evento%></td> 
</tr>
<tr>
	<th rowspan="2" colspan="2" width="86%" class="th2">Objetivos </th>
	<th colspan="2" width="14%" class="th2" align="center">Calificaci&oacute;n </th>
</tr>
<tr>
	<th width="7%" class="th2">Autoevaluador </th>
	<th width="7%" class="th2">Evaluador </th>
</tr>

<%	
'BUSCAR OBJETIVOS unicamente
i=0
Set l_rs = Server.CreateObject("ADODB.RecordSet")
l_sql = "SELECT DISTINCT evaobjetivo.evaobjnro, evaobjetivo.evaobjdext, evacab.evacabnro "
l_sql = l_sql & " FROM evaobjetivo "
l_sql = l_sql & " INNER JOIN evaluaobj    ON evaobjetivo.evaobjnro = evaluaobj.evaobjnro "
l_sql = l_sql & " INNER JOIN evadetevldor ON evadetevldor.evldrnro = evaluaobj.evldrnro "
l_sql = l_sql & " INNER JOIN evacab ON evacab.evacabnro = evadetevldor.evacabnro "
l_sql = l_sql & " WHERE evacab.empleado=" & l_ternro & " AND evacab.evaproynro="&l_evaproynro
rsOpen l_rs, cn, l_sql, 0
	
if not l_rs.eof then
	
	do while not l_rs.eof 
		i= i +1
%>
	<tr height="10">
		<td width="4%"><%=i%></td>
        <td valign=middle>	<%=trim(l_rs("evaobjdext"))%> </td>
<%			'Buscar RESULTADOS ASOCIADOS A LOS OBJS Y EVALUADORES. 
			Set l_rs2 = Server.CreateObject("ADODB.RecordSet")
			l_sql = "SELECT  Distinct evatipevalua.evatevdesabr, evadetevldor.evatevnro, evaluaobj.evaobjnro "
			l_sql = l_sql & " ,evaluaobj.evatrnro, evatipresu.evatrdesabr, evaobliorden  " 'evadetevldor.evldrnro, evadetevldor.evaseccnro
			l_sql = l_sql & " FROM evaluaobj "
			l_sql = l_sql & " INNER JOIN evaobjetivo ON evaobjetivo.evaobjnro = evaluaobj.evaobjnro "
			l_sql = l_sql & " INNER JOIN evadetevldor ON evadetevldor.evldrnro = evaluaobj.evldrnro "
			l_sql = l_sql & " INNER JOIN evasecc ON evasecc.evaseccnro = evadetevldor.evaseccnro "
			l_sql = l_sql & " INNER JOIN evatiposecc ON evatiposecc.tipsecnro = evasecc.tipsecnro "
			l_sql = l_sql & " INNER JOIN evaoblieva ON evaoblieva.evaseccnro = evadetevldor.evaseccnro AND evaoblieva.evatevnro= evadetevldor.evatevnro "
			l_sql = l_sql & " INNER JOIN evatipevalua ON evadetevldor.evatevnro = evatipevalua.evatevnro "
			l_sql = l_sql & " LEFT JOIN evatipresu ON evatipresu.evatrnro = evaluaobj.evatrnro " ' por si alguien no definio resultados.
			l_sql = l_sql & " WHERE evaluaobj.evaobjnro = " & l_rs("evaobjnro")
			l_sql = l_sql & "  AND  evasecc.tipsecnro =" &  cevaseccobj 
			l_sql = l_sql & "  AND (evadetevldor.evatevnro=" & cautoevaluador
			l_sql = l_sql & "       OR evadetevldor.evatevnro=" & cevaluador & ")" 
			l_sql = l_sql & "  AND evacabnro= " & l_rs("evacabnro") 
			l_sql = l_sql & "  ORDER BY evaoblieva.evaobliorden " 'evatipevalua.evatevdesabr
			rsOpen l_rs2, cn, l_sql, 0
			
			do while not l_rs2.eof 
%>
				<td align="center"> <%=l_rs2("evatrdesabr")%> </td>
<%
		   l_rs2.MoveNext
		   loop
		   l_rs2.Close
		   set l_rs2=nothing%>
	</tr>
<%
	l_rs.MoveNext
	loop
	
else  %>
   <tr height="20">
	 	<td colspan="8" align="center"> No existen Evaluaciones de Calific de Objetivos asociado al proyecto.</td> <!-- -->
   </tr>
<%
end if
  
l_rs.close
set l_rs=nothing
%>
</form>		
</table>

</body>
</html>

<%
cn.Close
Set cn = Nothing
%>

<% Option Explicit %>
<!--#include virtual="/ess/ess/shared/inc/const.inc"-->
<!--#include virtual="/serviciolocal/shared/db/conn_db.inc"-->
<!--#include virtual="/serviciolocal/shared/inc/const_eva.inc"-->
<% 
'=====================================================================================
'Archivo  : carga_retroalimentacion_COD_eva_00.asp
'Objetivo : ABM seguimiento x compromiso
'Fecha	  : 08-02-2005
'Autor	  : CCRossi
'Modificacion  : 21-03-2005 * cambiar letra en clase .rev
'            13-10-2005 - Leticia Amadio -  Adecuacion a Autogestion
'			 24/05/07 - Diego Rosso - Se agrego src="blanc.asp" para que funcione con https.
'=====================================================================================
 Dim l_rs
 Dim l_rs1
 Dim l_cm
 Dim l_sql
 Dim l_filtro
 Dim l_orden

'locales
 dim l_evacabnro 
 dim l_evatevnro 
 dim l_evaluador 
 dim l_planfecharev

 dim l_evatipobjnro 
 dim l_evaevenro 
 
'parametros
 Dim l_evldrnro
 
 l_evldrnro = request.querystring("evldrnro")

 if l_orden = "" then
  l_orden = " ORDER BY evaobjnro "
 end if

 Set l_rs = Server.CreateObject("ADODB.RecordSet")
 l_sql = "SELECT evaevenro, evatevnro, evaluador, evadetevldor.evacabnro  "
 l_sql = l_sql & " FROM  evadetevldor "
 l_sql = l_sql & " INNER JOIN empleado ON empleado.ternro = evadetevldor.evaluador "
 l_sql = l_sql & " INNER JOIN evacab ON evacab.evacabnro = evadetevldor.evacabnro "
 l_sql = l_sql & " WHERE evldrnro   = " & l_evldrnro
 rsOpen l_rs, cn, l_sql, 0
 if not l_rs.EOF then
	l_evaevenro = l_rs("evaevenro")
	l_evacabnro =l_rs("evacabnro")
	l_evatevnro =l_rs("evatevnro")
	l_evaluador =l_rs("evaluador")
 end if
 l_rs.close
 set l_rs=nothing

'Crear los evaplan de cada objetivo--------------------------------------------------

'busco el objetivo asociado al mismo evaluador, mismo evatevnro, misma cabecera.
Set l_rs = Server.CreateObject("ADODB.RecordSet")
l_sql = "SELECT evadetevldor.evldrnro, evaobjetivo.evaobjnro FROM evadetevldor "
l_sql = l_sql & " INNER JOIN evasecc ON evadetevldor.evaseccnro = evasecc.evaseccnro "
l_sql = l_sql & " INNER JOIN evatiposecc ON evasecc.tipsecnro = evatiposecc.tipsecnro "
l_sql = l_sql & " INNER JOIN evaluaobj   ON evaluaobj.evldrnro=evadetevldor.evldrnro "
l_sql = l_sql & "		 AND evaluaobj.evaborrador = 0 "
l_sql = l_sql & " INNER JOIN evaobjetivo ON evaobjetivo.evaobjnro=evaluaobj.evaobjnro "
l_sql = l_sql & " WHERE evacabnro = " & l_evacabnro
l_sql = l_sql & " AND   evatevnro = " & l_evatevnro
l_sql = l_sql & " AND   evaluador = " & l_evaluador
l_sql = l_sql & " AND   tipsecobj=-1" 
'Response.Write l_sql
rsOpen l_rs, cn, l_sql, 0
do while not l_rs.eof 
	Set l_rs1 = Server.CreateObject("ADODB.RecordSet")
	l_sql = "SELECT * FROM evaplan "
	l_sql = l_sql & " WHERE evaobjnro = " & l_rs("evaobjnro")
	l_sql = l_sql & " AND   evldrnro  = " & l_evldrnro
	'Response.Write l_sql
	rsOpen l_rs1, cn, l_sql, 0
	if  l_rs1.eof then
		l_rs1.Close
		set l_rs1=nothing
		l_sql= "insert into evaplan (evldrnro,evaobjnro) "
		l_sql = l_sql & " values (" & l_evldrnro & "," & l_rs("evaobjnro") &")"
'		Response.Write l_sql
		set l_cm = Server.CreateObject("ADODB.Command")  
		l_cm.activeconnection = Cn
		l_cm.CommandText = l_sql
		cmExecute l_cm, l_sql, 0
	else
		l_rs1.Close
		set l_rs1=nothing
	end if
	
	l_rs.MoveNext
loop	
l_rs.Close
set l_rs=nothing

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
    <tr>
        <th align=center class="th2">Compromiso</th>
        <th align=center class="th2">Observaciones del Seguimiento</th>
        <th align=center class="th2">Fecha de Reuni&oacute;n del Seguimiento</th>

    </tr>
<form name="datos" method="post">
<%
Set l_rs = Server.CreateObject("ADODB.RecordSet")
l_sql = "SELECT DISTINCT evaobjetivo.evaobjnro, evaobjetivo.evaobjdext,evaplan.aspectomejorar,"
l_sql = l_sql & " evaplan.planfecharev, evaplan.evaplnro, "
l_sql = l_sql & " evaobjetivo.evatipobjnro, evatipobjdabr ,evatipobjorden "
l_sql = l_sql & " ,evatipoobj.evatipobjorden, evatipopor "
l_sql = l_sql & " FROM evaplan "
l_sql = l_sql & " INNER JOIN evaobjetivo  ON evaobjetivo.evaobjnro = evaplan.evaobjnro"
l_sql = l_sql & " LEFT  JOIN evatipoobj ON evatipoobj.evatipobjnro = evaobjetivo.evatipobjnro"
l_sql = l_sql & " LEFT  JOIN evatipoobjpor ON evatipoobj.evatipobjnro = evatipoobjpor.evatipobjnro"
l_sql = l_sql & "		 AND evatipoobjpor.evaevenro = " & l_evaevenro
l_sql = l_sql & " WHERE evaplan.evldrnro =" & l_evldrnro
l_sql = l_sql & " ORDER BY evatipoobj.evatipobjorden "
'Response.Write l_sql
rsOpen l_rs, cn, l_sql, 0 
l_evatipobjnro=""
do until l_rs.eof
	if trim(l_rs("planfecharev"))="" or isnull(l_rs("planfecharev")) or l_rs("planfecharev")="null" then
		l_planfecharev = date()
	else	
		l_planfecharev = l_rs("planfecharev")
	end if	
		if l_evatipobjnro <> l_rs("evatipobjnro") then
		l_evatipobjnro= l_rs("evatipobjnro")  %>
		<tr>
        <td colspan="4"><b><%=l_rs("evatipobjdabr")%>
        <%if ccodelco=-1 then%>
			&nbsp;<%=l_rs("evatipopor")%>%
		<%end if%>	
        </b>
		</td>
		</tr>
	<%end if %>

<tr>
	<td align=center width=15%>
		<b><%=trim(l_rs("evaobjdext"))%></b>
	</td>
    <td align=center width=20%>
		<textarea class="rev" style="background : #e0e0de;" readonly name="aspectomejorar<%=l_rs("evaobjnro")%>"  cols=40 rows=4><%=trim(l_rs("aspectomejorar"))%></textarea>
	</td>
   <td nowrap width=10%>
		<input class="rev" style="background : #e0e0de;" readonly type="text" name="planfecharev<%=l_rs("evaobjnro")%>" size="10" value="<%=l_planfecharev%>">
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
<input type="Hidden" name="cabnro" value="0">
<iframe src="blanc.asp" name="grabar" style="visibility:hidden;width:0;height:0">
<!--iframe name="grabar" style="width:500;height:100"-->


</form>
</body>
</html>

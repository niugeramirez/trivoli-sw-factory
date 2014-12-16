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
'Autor			: 01  -06-2005
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
 
 dim l_proyectos 
 dim l_cantproys
 dim l_proys
 dim l_cantproysaprob
 dim l_sincerrarRDE
 
 dim l_ternro  
 dim l_evaevenro 
 dim l_area
  
' parametros de entrada---------------------------------------  
  l_evaevenro =  Request.QueryString("evaevenro") 
  l_ternro = Request.QueryString("ternro") 
  l_area = Request.QueryString("area") 
  
' ______________________________________________________________________________________________________________
' buscar todos los proyectos en que participo el empleado (igual periodo y estrnro que evento RDP - )
l_proyectos = "0" 
l_cantproys = 0 
l_proys="" 
l_cantproysaprob = 0 
Set l_rs1 = Server.CreateObject("ADODB.RecordSet")
l_sql = "SELECT DISTINCT evaproyecto.evaproynro, cabaprobada, evento.evaevenro "
l_sql = l_sql & " FROM evaevento "
l_sql = l_sql & " INNER JOIN evaproyecto ON evaproyecto.evapernro = evaevento.evaperact "
l_sql = l_sql & " INNER JOIN evaevento evento ON evento.evaproynro = evaproyecto.evaproynro "
l_sql = l_sql & " INNER JOIN evacab ON evacab.evaproynro = evaproyecto.evaproynro "
l_sql = l_sql & " INNER JOIN evatipoeva ON evatipoeva.evatipnro = evaevento.evatipnro "
l_sql = l_sql & " INNER JOIN evatip_estr ON evatip_estr.evatipnro = evatipoeva.evatipnro AND evatip_estr.estrnro=evaproyecto.estrnro  AND evatip_estr.tenro ="& cdepartamento 
l_sql = l_sql & " WHERE  evaevento.evaevenro ="& l_evaevenro &" AND evacab.empleado="&l_ternro
		' evacab.cabaprobada= -1 AND 	

rsOpen l_rs1, cn, l_sql, 0 
do while not l_rs1.eof 
	if l_rs1("cabaprobada") = -1 then 
		l_proyectos = l_proyectos & "," & l_rs1("evaproynro") 
		l_cantproysaprob = l_cantproysaprob + 1 
	else 
		l_proys = l_proys &  " - " & l_rs1("evaevenro")  'l_rs1("evaproynro")
	end if 
	l_cantproys = l_cantproys +1 
l_rs1.MoveNext 
loop 
l_rs1.close 

l_proyectos = Split(l_proyectos,",") 

l_sincerrarRDE="NO" 
if l_cantproys <> l_cantproysaprob then 
	l_sincerrarRDE="SI" 
end if 

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

l_sql= " SELECT terape, terape2,ternom, ternom2 "
l_sql = l_sql & " FROM tercero  WHERE ternro= " & l_ternro
		rsOpen l_rs1, cn, l_sql, 0 
l_empleado = l_rs1("terape") & " " &  l_rs1("terape2") & " " & l_rs1("ternom") &  " "  & l_rs1("ternom2")
l_rs1.Close
%>

<table border="0" cellpadding="0" cellspacing="1" width="100%">
<% if l_sincerrarRDE="SI" then %>
<tr height="20">
	<td colspan="4" align="left" width="25%">
		<b> AVISO:</b> No se detallan los comentarios de Areas que no tienen sus RDE's cerradas. <br>
		  &nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp; &nbsp; Eventos sin RDE's cerradas: <%=l_proys%>
	</td>
</tr>
<tr><td colspan="4" align="center">&nbsp;</td> </tr>
<% end if %>
<tr height="20">
	<td colspan="4" align="center"><b>Comentarios de Areas RDE </b></td> 
</tr>
<tr height="20">
	<td colspan="4"><b>Empleado</b>: <%=l_empleado%></td> 
</tr>
<tr height="20">
	<td colspan="4">&nbsp;</td> 
</tr>

<%
Set l_rs = Server.CreateObject("ADODB.RecordSet")
l_sql = " SELECT evatitdesabr FROM evatitulo  WHERE evatitnro=" & l_area
rsOpen l_rs, cn, l_sql, 0

if not l_rs.eof then %>
<tr>
	<th colspan="4" align="left" class="th2"><%= UCase(l_rs("evatitdesabr"))%></th>
</tr>

<% if UBound(l_proyectos)= 0 then %>
	<tr>
		<td colspan="4" align="left"> No existen proyectos con RDE's cerradas para el area</td>
	</tr>
<%	else
		
		for i = 1 to ubound(l_proyectos) 
			
			l_sql = " SELECT DISTINCT evatipevalua.evatevnro, evatevdesabr, evaareacom, evaevenro " '
			l_sql = l_sql & " FROM evacab "
			l_sql = l_sql & " INNER JOIN evadetevldor ON evadetevldor.evacabnro =evacab.evacabnro " 
			l_sql = l_sql & " LEFT JOIN  evaareacom  ON  evaareacom.evldrnro= evadetevldor.evldrnro" 
			l_sql = l_Sql & " INNER JOIN evatipevalua ON evatipevalua.evatevnro = evadetevldor.evatevnro "
            l_sql = l_sql & " WHERE evaproynro="& l_proyectos(i)  & " AND evacab.empleado="& l_ternro & " AND evaareacom.evatitnro = " & l_area 
			l_sql = l_sql & " ORDER BY evatipevalua.evatevnro " 'evaoblieva.evaobliorden "
			'response.write l_sql
			rsOpen l_rs1, cn, l_sql, 0 %>
			<tr>
				<td align="left" width="10%" nowrap><strong>Evento nro.</strong> </td>
				<td colspan="3" width="90%" align="left"> <%= l_rs1("evaevenro")%></td> <!--  'l_proyectos(i)-->
			</tr>	
<%			do while not l_rs1.eof %>
				<tr>
					<td align="left"><%=l_rs1("evatevdesabr")%></td>
					<td colspan="3" align="left"><%=unescape(l_rs1("evaareacom"))%></td>
				</tr>	
<%			l_rs1.MoveNext 
			loop 
			l_rs1.Close 
		next 
		
		
	end if
	
else ' no existe el area???? %>
<tr>
	<td colspan="4" align="left"> No existe el Area. </td>
</tr>
<%end if 

l_rs.close
set l_rs=nothing

'l_sql = l_sql & " INNER JOIN evasecc ON evasecc.evaseccnro = evadetevldor.evaseccnro "
'l_sql = l_sql & " INNER JOIN evatiposecc ON evatiposecc.tipsecnro = evasecc.tipsecnro "
'l_sql = l_sql & " INNER JOIN evaoblieva ON evaoblieva.evaseccnro = evadetevldor.evaseccnro AND evaoblieva.evatevnro= evadetevldor.evatevnro "
'l_sql = l_sql & " INNER JOIN evatipevalua ON evadetevldor.evatevnro = evatipevalua.evatevnro "
'l_sql = l_sql & " LEFT  JOIN evatipresu ON evatipresu.evatrnro = evaluaobj.evatrnro " ' por si alguien no definio resultados.
'l_sql = l_sql & " ORDER BY evaoblieva.evaobliorden "
%>
</form>		
</table>

</body>
</html>

<%
cn.Close
Set cn = Nothing
%>

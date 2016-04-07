<% Option Explicit
if request.querystring("excel") then
	Response.AddHeader "Content-Disposition", "attachment;filename=Pagos entre Fechas.xls" 
	Response.ContentType = "application/vnd.ms-excel"
end if
 %>
<!--#include virtual="/turnos/shared/db/conn_db.inc"-->
<!--#include virtual="/turnos/shared/inc/fecha.inc"-->
<% 

Dim l_rs
Dim l_sql
Dim l_filtro
Dim l_orden
Dim l_sqlfiltro
Dim l_sqlorden
Dim l_totvol
Dim l_cant
dim l_fechahorainicio
dim l_cantturnossimult
dim l_idmedio
dim l_cantturnos
dim l_fondo

Dim l_primero
Dim l_fechadesde
Dim l_fechahasta
Dim l_descripcion
Dim l_titulo
Dim l_medico

Dim l_can_calendarios
Dim l_can_turnos
Dim l_can_visitas

Dim l_idrecursoreservable

l_filtro = request("filtro")
l_orden  = request("orden")

if l_orden = "" then
  l_orden = " ORDER BY pagos.idmediodepago , obrassociales.descripcion , pagos.fecha "
end if

sub encabezado
 %>
 <!--
	<tr>
        <td  colspan="3" align="center" ><h3>Medio de Pago:&nbsp;<%'= l_rs("titulo") %></h3></td>	
		<td  colspan="3" align="center" ><h3>Medico:&nbsp;<%'= l_medico %></h3></td>	
    </tr>	-->

    <tr>
        <th width="100">Calendarios</th>
        <th width="200">Turnos</th>	
        <th width="200">Visitas con Turno</th>			
		<!--<th width="200">Visitas sin Turno</th> -->

	
    </tr>
<%
end sub	


%>
%>
<!DOCTYPE HTML PUBLIC "-//IETF//DTD HTML//EN">
<html>
<script src="/turnos/shared/js/fn_windows.js"></script>
<script src="/turnos/shared/js/fn_confirm.js"></script>
<script src="/turnos/shared/js/fn_ayuda.js"></script>
<SCRIPT LANGUAGE="Javascript" SRC="../FusionCharts/FusionCharts.js"></SCRIPT>
<head>
<% if request.querystring("excel") = false then  %>
<link href="/turnos/ess/shared/css/tables_gray.css" rel="StyleSheet" type="text/css">
<% End If %>


<%
	'We've included ../Includes/FusionCharts.asp, which contains functions
	'to help us easily embed the charts.
	%>
<!-- #INCLUDE virtual="/turnos/Includes/FusionCharts.asp" -->
	<%
	'We've also included ../Includes/FC_Colors.asp, having a list of colors
	'to apply different colors to the chart's columns. We provide a function for it - getFCColor()
	%>
<!-- #INCLUDE virtual="/turnos/Includes/FC_Colors.asp" -->


<meta http-equiv="Content-Type" content="text/html; charset=iso-8859-1">
<title>Estadisticas entre Fechas</title>
</head>

<script>
var jsSelRow = null;

function Deseleccionar(fila){
	fila.className = "MouseOutRow";
}

function Seleccionar(fila,cabnro, turnoid){
	if (jsSelRow != null){
		Deseleccionar(jsSelRow);
	};
	document.datos.cabnro.value = cabnro;
	document.datos.idturno.value = turnoid;
	fila.className = "SelectedRow";
	jsSelRow = fila;
}
</script>
<% 


l_filtro = replace (l_filtro, "*", "%")
l_idmedio = request("idmedio")
l_fechadesde = request("qfechadesde")
l_fechahasta = request("qfechahasta")
l_idrecursoreservable = request("idrecursoreservable")

Set l_rs = Server.CreateObject("ADODB.RecordSet")

' Obtengo el Nombre del Medio de Pago
'if l_idmedio = "0" then
'	l_titulo = "Todos"
'else	
'	l_sql = "SELECT  * "
'	l_sql = l_sql & " FROM mediosdepago "
'	l_sql = l_sql & " WHERE id = " & l_idmedio
'	rsOpen l_rs, cn, l_sql, 0 
'	l_titulo = ""
'	if not l_rs.eof then
'		l_titulo = l_rs("titulo")
'	end if
'	l_rs.close
'end if

' Obtengo el Nombre del Medico
if l_idrecursoreservable = "0" then
	l_medico = "Todos"
else	
	l_sql = "SELECT  * "
	l_sql = l_sql & " FROM recursosreservables "
	l_sql = l_sql & " WHERE id = " & l_idrecursoreservable
	rsOpen l_rs, cn, l_sql, 0 
	l_medico = ""
	if not l_rs.eof then
		l_medico = l_rs("descripcion")
	end if
	l_rs.close
end if


' Cantidad de Calendarios disponibles
l_sql = "SELECT  count(*) "
l_sql = l_sql & " FROM calendarios "
l_sql = l_sql & " WHERE  CONVERT(VARCHAR(10), calendarios.fechahorainicio, 101) >= " & cambiafecha(l_fechadesde,"YMD",true) 
l_sql = l_sql & " AND    CONVERT(VARCHAR(10), calendarios.fechahorainicio, 101) <= " & cambiafecha(l_fechahasta,"YMD",true) 
if l_idrecursoreservable <> "0" then
	l_sql = l_sql & " AND calendarios.idrecursoreservable = " & l_idrecursoreservable
end if	
l_sql = l_sql & " and calendarios.empnro = " & Session("empnro")   
'response.write l_sql & "<br>"
rsOpen l_rs, cn, l_sql, 0 
'response.write l_rs(0) & "<br>"
l_can_calendarios = l_rs(0)
l_rs.Close


' Turnos Sacados asociados a los Calendarios
l_sql = "SELECT  count(*) "
l_sql = l_sql & " FROM calendarios "
l_sql = l_sql & " INNER JOIN turnos ON turnos.idcalendario = calendarios.id "
l_sql = l_sql & " WHERE  CONVERT(VARCHAR(10), calendarios.fechahorainicio, 101)  >= " & cambiafecha(l_fechadesde,"YMD",true) 
l_sql = l_sql & " AND  CONVERT(VARCHAR(10), calendarios.fechahorainicio, 101) <= " & cambiafecha(l_fechahasta,"YMD",true) 
if l_idrecursoreservable <> "0" then
	l_sql = l_sql & " AND calendarios.idrecursoreservable = " & l_idrecursoreservable
end if	
l_sql = l_sql & " and calendarios.empnro = " & Session("empnro")   
rsOpen l_rs, cn, l_sql, 0 
'response.write l_rs(0) & "<br>"
l_can_turnos = l_rs(0)
l_rs.Close

' Visitas de los Turnos
l_sql = "SELECT  count(*) "
l_sql = l_sql & " FROM visitas "
l_sql = l_sql & " INNER JOIN turnos ON turnos.id = visitas.idturno "
l_sql = l_sql & " INNER JOIN calendarios ON calendarios.id = turnos.idcalendario "
l_sql = l_sql & " WHERE  CONVERT(VARCHAR(10), calendarios.fechahorainicio, 101)  >= " & cambiafecha(l_fechadesde,"YMD",true) 
l_sql = l_sql & " AND  CONVERT(VARCHAR(10), calendarios.fechahorainicio, 101) <= " & cambiafecha(l_fechahasta,"YMD",true) 
if l_idrecursoreservable <> "0" then
	l_sql = l_sql & " AND calendarios.idrecursoreservable = " & l_idrecursoreservable
end if	
l_sql = l_sql & " and visitas.empnro = " & Session("empnro")   
'response.write l_sql & "<br>"
rsOpen l_rs, cn, l_sql, 0 
'response.write l_rs(0) & "<br>"
l_can_visitas = l_rs(0)
l_rs.Close




 %>

<body leftmargin="0" rightmargin="0" topmargin="0" bottommargin="0" onload="//javascript:parent.Buscar();">
<table>
    <tr>
        <td colspan="6">&nbsp;</td>
    </tr>
	<tr>
        <td  colspan="6" align="center" ><h3>Estadisticas de Cantidades desde:&nbsp;<%= l_fechadesde %>&nbsp; al <%= l_fechahasta %>&nbsp;&nbsp;</h3></td>	
    </tr>

<% 	

	encabezado
	
    
	%>
	    <tr>
			
	        <td align="center"><%= l_can_calendarios %></td>	
			<td align="center"><%= l_can_turnos%></td>
			<td align="center"><%= l_can_visitas %></td>				
				
										   
	    </tr>
		
		 <tr>
		 <td align="center" colspan=15>
	<%



	Dim arrData(3,2)
	
	'Dim l_rs
	'Dim l_sql
	
	Dim l_cadena
	Dim l_lista
	Dim l_elem
	Dim i
	
	l_cadena = "Calendario-" & l_can_calendarios & "@" & "Turnos-" & l_can_turnos & "@" & "Visitas-" & l_can_visitas & "@"
	
	l_lista = Split(l_cadena,"@")
	
	i = 0
	
	do while i <= UBound(l_lista)-1

		l_elem = Split(l_lista(i) , "-")

		arrData(i,1) = l_elem(0)
		arrData(i,2) = l_elem(1)
		
		i = i + 1

	loop


	'Now, we need to convert this data into XML. We convert using string concatenation.
	Dim strXML
	'Initialize <graph> element
	'strXML = "<graph caption='Estadísticas de Pedidos' numberPrefix='$' formatNumberScale='0' decimalPrecision='0'>"
	strXML = "<graph caption='' subCaption='' yaxisname='Cantidad' xaxisname='' formatNumberScale='0' decimalPrecision='2' showNames='1' showValues='1' showPercentageInLabel ='1' showAlternateVGridColor='1' alternateVGridAlpha='10' alternateVGridColor='AFD8F8' numDivLines='4' decimalPrecision='0' canvasBorderThickness='1' canvasBorderColor='114B78' baseFontColor='114B78' hoverCapBorderColor='114B78' hoverCapBgColor='E7EFF6'> "
	
	'Convert data to XML and append
	For i=0 to UBound(arrData)-1
		'add values using <set name='...' value='...' color='...'/>
		strXML = strXML & "<set name='" & arrData(i,1) & "' value='" & replace(arrData(i,2),",",".") & "' color='" & getFCColor() & "' />"
	Next
	'Close <graph> element
	strXML = strXML & "</graph>"

	'response.write strXML & "-"
	'response.end
	
	
'	'Create the chart - Column 3D Chart with data contained in strXML
'		Call renderChart("../../FusionCharts/FCF_Column3D.swf", "", strXML, "productSales", 600, 400)
		Call renderChart("../FusionCharts/FCF_Bar2D.swf", "", strXML, "productSales", 600, 250)
	
	set l_rs = Nothing
	cn.Close
	set cn = Nothing	
%>

	</td>	
<BR><BR>
		
	<%	
	'totales


l_rs.Close
set l_rs = Nothing
cn.Close
set cn = Nothing
%>

</table>
<form name="datos" method="post">
<input type="hidden" name="cabnro" value="0">
<input type="hidden" name="idturno" value="0">
<input type="hidden" name="orden" value="<%= l_orden %>">
<input type="hidden" name="filtro" value="<%= l_filtro %>">
</form>
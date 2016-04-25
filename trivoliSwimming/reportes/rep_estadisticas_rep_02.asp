<% Option Explicit
if request.querystring("excel") then
	Response.AddHeader "Content-Disposition", "attachment;filename=Pagos entre Fechas.xls" 
	Response.ContentType = "application/vnd.ms-excel"
end if
 %>
<!--#include virtual="/trivoliSwimming/shared/db/conn_db.inc"-->
<!--#include virtual="/trivoliSwimming/shared/inc/fecha.inc"-->
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
        <th width="100">Articulo</th>
        <th width="200">Ventas Proyectadas</th>	
        <th width="200">Ventas Reales</th>			
		<!--<th width="200">Visitas sin Turno</th> -->

	
    </tr>
<%
end sub	


sub encabezado2
 %>

    <tr>
        <th width="100">Articulo</th>
		<th width="200">Cantidad Compras</th>	
		<th width="200">Cantidad Venta</th>	
        <th width="200">Stock</th>	
        <!--<th width="200">Cantidad Venta</th>			
		<th width="200">Visitas sin Turno</th> -->

	
    </tr>
<%
end sub	

sub encabezado3
 %>


    <tr>
        <th width="100">Estado Instalacion</th>
        <th width="200">Cantidad</th>			

	
    </tr>
<%
end sub	

sub encabezado4
 %>


    <tr>
        <th width="100">Medio de Pago</th>
        <th width="200">Entradas</th>			
		<th width="200">Salidas</th>	
		<th width="200">Saldo</th>					

	
    </tr>
<%
end sub

sub encabezado5
 %>


    <tr>
        <th width="100">Reponsable</th>
		<th width="100">Medio de Pago</th>
        <th width="200">Entradas</th>			
		<th width="200">Salidas</th>	
		<th width="200">Saldo</th>					

	
    </tr>
<%
end sub	
%>
<!DOCTYPE HTML PUBLIC "-//IETF//DTD HTML//EN">
<html>
<script src="/trivoliSwimming/shared/js/fn_windows.js"></script>
<script src="/trivoliSwimming/shared/js/fn_confirm.js"></script>
<script src="/trivoliSwimming/shared/js/fn_ayuda.js"></script>
<SCRIPT LANGUAGE="Javascript" SRC="../FusionCharts/FusionCharts.js"></SCRIPT>
<head>
<% if request.querystring("excel") = false then  %>
<link href="/trivoliSwimming/ess/shared/css/tables_gray.css" rel="StyleSheet" type="text/css">
<% End If %>



<%


	'We've included ../Includes/FusionCharts.asp, which contains functions
	'to help us easily embed the charts.
	%>
<!-- #INCLUDE virtual="/trivoliSwimming/Includes/FusionCharts.asp" -->
	<%
	'We've also included ../Includes/FC_Colors.asp, having a list of colors
	'to apply different colors to the chart's columns. We provide a function for it - getFCColor()
	%>
<!-- #INCLUDE virtual="/trivoliSwimming/Includes/FC_Colors.asp" -->


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

	Dim arrData(5,5)
	
	'Dim l_rs
	'Dim l_sql
	
	Dim l_cadena
	Dim l_lista
	Dim l_elem
	Dim i


l_filtro = replace (l_filtro, "*", "%")
l_idmedio = request("idmedio")
l_fechadesde = request("qfechadesde")
l_fechahasta = request("qfechahasta")
l_idrecursoreservable = request("idrecursoreservable")

Set l_rs = Server.CreateObject("ADODB.RecordSet")



select case l_idrecursoreservable

case 1

l_sql = " select articulo_id " 
l_sql = l_sql & "		,articulo_desc "
l_sql = l_sql & "		,SUM(cant_vtas_proyectadas)												as cant_vtas_proyectadas	"
l_sql = l_sql & "		,SUM(cant_vtas_reales)													as cant_vtas_reales "
l_sql = l_sql & " from ( "
l_sql = l_sql & "		SELECT  proyeccionventas.idconceptoCompraVenta													as articulo_id "
l_sql = l_sql & "				,conceptosCompraVenta.descripcion														as articulo_desc "
l_sql = l_sql & "				,(proyeccionventas.cantidadproyectada)												as cant_vtas_proyectadas	"
l_sql = l_sql & "				,(	select sum(detalleVentas.cantidad) "
l_sql = l_sql & "					from detalleVentas "
l_sql = l_sql & "					inner join ventas on ventas.id = detalleVentas.idventa "
l_sql = l_sql & "					where detalleVentas.idconceptoCompraVenta = proyeccionventas.idconceptoCompraVenta "
l_sql = l_sql & "					and ventas.fecha >= proyeccionventas.fecha_desde "
l_sql = l_sql & "					and ventas.fecha <= proyeccionventas.fecha_hasta "
l_sql = l_sql & "				 )																						as cant_vtas_reales "
l_sql = l_sql & "		FROM proyeccionventas "
l_sql = l_sql & "		inner join conceptosCompraVenta on conceptosCompraVenta.id = proyeccionventas.idconceptoCompraVenta "
'l_sql = l_sql & "		WHERE CONVERT(VARCHAR(10), proyeccionventas.fecha_desde, 111) >= " & cambiafecha(l_fechadesde,"YMD",true)  
'l_sql = l_sql & "		AND CONVERT(VARCHAR(10), proyeccionventas.fecha_desde, 111) <= " & cambiafecha(l_fechahasta,"YMD",true) 
l_sql = l_sql & "		WHERE proyeccionventas.fecha_desde >= " & cambiafecha(l_fechadesde,"YMD",true)  
l_sql = l_sql & "		AND proyeccionventas.fecha_desde <= " & cambiafecha(l_fechahasta,"YMD",true) 
l_sql = l_sql & "		AND proyeccionventas.empnro =  " & Session("empnro")  
l_sql = l_sql & "	) tab_agrup "
l_sql = l_sql & " group by articulo_id "
l_sql = l_sql & "		,articulo_desc	"
l_sql = l_sql & " ORDER BY articulo_desc "
rsOpen l_rs, cn, l_sql, 0 

'response.write l_sql&"</br>"

 %>

<body leftmargin="0" rightmargin="0" topmargin="0" bottommargin="0" onload="//javascript:parent.Buscar();">
<table>
    <tr>
        <td colspan="6">&nbsp;</td>
    </tr>
	<tr>
        <td  colspan="6" align="center" ><h3>Ventas desde:&nbsp;<%= l_fechadesde %>&nbsp; al <%= l_fechahasta %>&nbsp;&nbsp;</h3></td>	
    </tr>
<% 	
	encabezado

	i = 0
    do while not l_rs.eof
	
	    arrData(i,1) = l_rs("articulo_desc")
		arrData(i,2) = l_rs("cant_vtas_proyectadas")
		arrData(i,3) = l_rs("cant_vtas_reales")
	%>
	    <tr>			
	        <td align="center"><%= l_rs("articulo_desc") %></td>	
			<td align="center"><%= l_rs("cant_vtas_proyectadas")%></td>
			<td align="center"><%= l_rs("cant_vtas_reales") %></td>												   
	    </tr>		
		 <tr>
		 <td align="center" colspan=15>
	<%
		i = i + 1	
		l_rs.movenext
	loop

	'Now, we need to convert this data into XML. We convert using string concatenation.
	Dim strXML
	Dim strXML_C
	Dim strXML_P
	Dim strXML_R
	'Initialize <graph> element
	'strXML = "<graph caption='Estadísticas de Pedidos' numberPrefix='$' formatNumberScale='0' decimalPrecision='0'>"
	'strXML = "<graph caption='' subCaption='' yaxisname='Cantidad' xaxisname='' formatNumberScale='0' decimalPrecision='2' showNames='1' showValues='1' showPercentageInLabel ='1' showAlternateVGridColor='1' alternateVGridAlpha='10' alternateVGridColor='AFD8F8' numDivLines='4' decimalPrecision='0' canvasBorderThickness='1' canvasBorderColor='114B78' baseFontColor='114B78' hoverCapBorderColor='114B78' hoverCapBgColor='E7EFF6'> "
	
	strXML = "<graph xaxisname='Articulos' yaxisname='Cantidad' hovercapbg='DEDEBE' hovercapborder='889E6D' rotateNames='0' animation='1' yAxisMaxValue='100' numdivlines='9' divLineColor='CCCCCC' divLineAlpha='80' decimalPrecision='0' showAlternateVGridColor='1' AlternateVGridAlpha='30' AlternateVGridColor='CCCCCC' caption='Ventas' subcaption='Proyectadas Vs Reales' > "
	
	'Convert data to XML and append
	For i=0 to UBound(arrData)-1
		'add values using <set name='...' value='...' color='...'/>
		'strXML = strXML & "<set name='" & arrData(i,1) & "' value='" & replace(arrData(i,2),",",".") & "' color='" & getFCColor() & "' />"
		'strXML = strXML & "<set name='" & arrData(i,1) & "' value='" & replace(arrData(i,3),",",".") & "' color='" & getFCColor() & "' />"
   		strXML_C = strXML_C & "   <category name='" & arrData(i,1) & "' hoverText=' "& arrData(i,1) &"'/> "		
		strXML_P = strXML_P & "    <set value='" & arrData(i,2) & "' /> "
		strXML_R = strXML_R & "    <set value='" & arrData(i,3) & "' /> "
	Next

   strXML = strXML & "<categories font='tahoma' fontSize='11' fontColor='000000'> "	
   strXML = strXML & strXML_C
   strXML = strXML & "</categories> "
   strXML = strXML & " <dataset seriesname='Proyectado' color='FDC12E' alpha='70'> " 
   strXML = strXML & strXML_P
   strXML = strXML & " </dataset> "
   strXML = strXML & " <dataset seriesname='Real' color='56B9F9' showValues='1' alpha='70'> "
   strXML = strXML & strXML_R
   strXML = strXML & " </dataset> "
   strXML = strXML & " </graph>" 
	
	'Create the chart - Column 3D Chart with data contained in strXML
	Call renderChart("../FusionCharts/FCF_MSColumn3D.swf", "", strXML, "productSales", 800, 350)
	
	set l_rs = Nothing
	cn.Close
	set cn = Nothing	
		
%>

	</td>	
	</table>
<BR><BR>
		
<%	
	
case 2

reDim arrData(7,7)

l_sql = " SELECT conceptosCompraVenta.id "
l_sql = l_sql & " ,conceptosCompraVenta.descripcion as articulo "
l_sql = l_sql & " ,( select ISNULL(SUM( detalleCompras.cantidad) ,0) "
l_sql = l_sql & " from detalleCompras "
l_sql = l_sql & " inner join compras on compras.id = detalleCompras.idcompra "
l_sql = l_sql & " where detalleCompras.idconceptoCompraVenta = conceptosCompraVenta.id "
l_sql = l_sql & " AND compras.fecha >=  " & cambiafecha(l_fechadesde,"YMD",true) 
l_sql = l_sql & " AND compras.fecha <=  " & cambiafecha(l_fechahasta ,"YMD",true) 
			
l_sql = l_sql & " ) as cantidad_compra "
l_sql = l_sql & " ,( select ISNULL(SUM( detalleVentas.cantidad),0) "
l_sql = l_sql & " from detalleVentas "
l_sql = l_sql & " inner join ventas on ventas.id = detalleVentas.idVenta "
l_sql = l_sql & " where detalleVentas.idconceptoCompraVenta = conceptosCompraVenta.id "


l_sql = l_sql & " AND ventas.fecha >=  " & cambiafecha(l_fechadesde,"YMD",true) 
l_sql = l_sql & " AND ventas.fecha <=  " & cambiafecha(l_fechahasta ,"YMD",true) 
		
l_sql = l_sql & " )	as cantidad_venta "
l_sql = l_sql & " ,( select ISNULL(SUM( detalleCompras.cantidad),0) "
l_sql = l_sql & " from detalleCompras "
l_sql = l_sql & " inner join compras on  compras.id = detalleCompras.idcompra " 
l_sql = l_sql & " where detalleCompras.idconceptoCompraVenta = conceptosCompraVenta.id "
l_sql = l_sql & " AND compras.fecha >=  " & cambiafecha(l_fechadesde,"YMD",true) 
l_sql = l_sql & " AND compras.fecha <=  " & cambiafecha(l_fechahasta ,"YMD",true) 		
l_sql = l_sql & " ) - "
l_sql = l_sql & " (select ISNULL(SUM( detalleVentas.cantidad),0) "
l_sql = l_sql & " from detalleVentas " 
l_sql = l_sql & " inner join ventas on ventas.id = detalleVentas.idVenta "
l_sql = l_sql & " where detalleVentas.idconceptoCompraVenta = conceptosCompraVenta.id "
l_sql = l_sql & " AND ventas.fecha >=  " & cambiafecha(l_fechadesde,"YMD",true) 
l_sql = l_sql & " AND ventas.fecha <=  " & cambiafecha(l_fechahasta ,"YMD",true) 		
l_sql = l_sql & " ) as stock "
l_sql = l_sql & " FROM conceptosCompraVenta "
l_sql = l_sql & " where conceptosCompraVenta.empnro =  " & Session("empnro")
l_sql = l_sql & " group by conceptosCompraVenta.id "
l_sql = l_sql & "  ,conceptosCompraVenta.descripcion "
l_sql = l_sql & " having ( select ISNULL(SUM( detalleCompras.cantidad),0) "
l_sql = l_sql & " from detalleCompras "
l_sql = l_sql & " inner join compras on compras.id = detalleCompras.idcompra "
l_sql = l_sql & " where detalleCompras.idconceptoCompraVenta = conceptosCompraVenta.id "
l_sql = l_sql & " AND compras.fecha >=  " & cambiafecha(l_fechadesde,"YMD",true) 
l_sql = l_sql & " AND compras.fecha <=  " & cambiafecha(l_fechahasta ,"YMD",true) 					
l_sql = l_sql & " ) - "
l_sql = l_sql & " ( select ISNULL(SUM( detalleVentas.cantidad),0) "
l_sql = l_sql & " from detalleVentas "
l_sql = l_sql & " inner join ventas on ventas.id = detalleVentas.idVenta "
l_sql = l_sql & " where detalleVentas.idconceptoCompraVenta = conceptosCompraVenta.id "
l_sql = l_sql & " AND ventas.fecha >=  " & cambiafecha(l_fechadesde,"YMD",true) 
l_sql = l_sql & " AND ventas.fecha <=  " & cambiafecha(l_fechahasta ,"YMD",true) 		
l_sql = l_sql & " ) <> 0 ORDER BY conceptosCompraVenta.descripcion "

'response.write l_sql&"</br>"
rsOpen l_rs, cn, l_sql, 0 


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

	encabezado2


		
	
		i = 0
    do while not l_rs.eof
	
		'response.write l_rs(1) & "<br>"
		'response.write l_rs(2) & "<br>"
		'response.write l_rs(3) & "<br>"
		'response.end
	
	    arrData(i,1) = l_rs("articulo")
		arrData(i,2) = l_rs("cantidad_compra")
		arrData(i,3) = l_rs("cantidad_venta")
		arrData(i,4) = l_rs("stock")
	%>
	    <tr>
			
	        <td align="center"><%= l_rs("articulo") %></td>	
			<td align="center"><%= l_rs("cantidad_compra")%></td>
			<td align="center"><%= l_rs("cantidad_venta")%></td>
			<td align="center"><%= l_rs("stock")%></td>
													   
	    </tr>
		
		 <tr>
		 <td align="center" colspan=15>
	<%
		i = i + 1	
		l_rs.movenext
	loop
	

	'Now, we need to convert this data into XML. We convert using string concatenation.
	Dim strXML1
	Dim strXML_C1
	Dim strXML_1P
	Dim strXML_R1
	Dim strXML_s1
	'Initialize <graph> element
	'strXML = "<graph caption='Estadísticas de Pedidos' numberPrefix='$' formatNumberScale='0' decimalPrecision='0'>"
	strXML = "<graph caption='' subCaption='' yaxisname='Cantidad' xaxisname='' formatNumberScale='0' decimalPrecision='2' showNames='1' showValues='1' showPercentageInLabel ='1' showAlternateVGridColor='1' alternateVGridAlpha='10' alternateVGridColor='AFD8F8' numDivLines='4' decimalPrecision='0' canvasBorderThickness='1' canvasBorderColor='114B78' baseFontColor='114B78' hoverCapBorderColor='114B78' hoverCapBgColor='E7EFF6'> "
	
	strXML = "<graph xaxisname='Articulos' yaxisname='Cantidad' hovercapbg='DEDEBE' hovercapborder='889E6D' rotateNames='0' animation='1' yAxisMaxValue='100' numdivlines='9' divLineColor='CCCCCC' divLineAlpha='80' decimalPrecision='0' showAlternateVGridColor='1' AlternateVGridAlpha='30' AlternateVGridColor='CCCCCC' caption='Stock' subcaption='' > "

	
	'Convert data to XML and append
	For i=0 to UBound(arrData)-1
		'add values using <set name='...' value='...' color='...'/>
		'strXML = strXML & "<set name='" & arrData(i,1) & "' value='" & replace(arrData(i,2),",",".") & "' color='" & getFCColor() & "' />"
		'strXML = strXML & "<set name='" & arrData(i,1) & "' value='" & replace(arrData(i,3),",",".") & "' color='" & getFCColor() & "' />"
   		strXML_C1 = strXML_C1 & "   <category name='" & arrData(i,1) & "' hoverText=' "& arrData(i,1) &"'/> "		
		strXML_1P = strXML_1P & "    <set value='" & arrData(i,2) & "' /> "
		strXML_R1 = strXML_R1 & "    <set value='" & arrData(i,3) & "' /> "
		strXML_s1 = strXML_s1 & "    <set value='" & arrData(i,4) & "' /> "
	Next
	'Close <graph> element
	'strXML = strXML & "</graph>"

	'response.write strXML & "-"
	'response.end
	
	
	'strXML = "<graph xaxisname='Continent' yaxisname='Export' hovercapbg='DEDEBE' hovercapborder='889E6D' rotateNames='0' animation='1' yAxisMaxValue='100' numdivlines='9' divLineColor='CCCCCC' divLineAlpha='80' decimalPrecision='0' showAlternateVGridColor='1' AlternateVGridAlpha='30' AlternateVGridColor='CCCCCC' caption='Global Export' subcaption='In Millions Tonnes per annum pr Hectare' > "
   'strXML = strXML & "<categories font='Arial' fontSize='11' fontColor='000000'> "
   'strXML = strXML & "   <category name='N. America' hoverText='North America'/> "
   'strXML = strXML & "   <category name='Asia' /> "
   'strXML = strXML & "   <category name='Europe' /> "
   'strXML = strXML & "   <category name='Australia' /> "
   'strXML = strXML & "   <category name='Africa' /> "
   strXML = strXML & "<categories font='tahoma' fontSize='11' fontColor='000000'> "	
   strXML = strXML & strXML_C1
   strXML = strXML & "</categories> "
  strXML = strXML & " <dataset seriesname='Compra' color='FDC12E' alpha='70'> " 
    strXML = strXML & strXML_1P
  strXML = strXML & " </dataset> "
  strXML = strXML & " <dataset seriesname='Venta' color='8E468E' showValues='1' alpha='70'> "
  strXML = strXML & strXML_R1
  strXML = strXML & " </dataset> "
  strXML = strXML & " <dataset seriesname='Stock' color='B3AA00' showValues='1' alpha='70'> "
  strXML = strXML & strXML_s1
  strXML = strXML & " </dataset> "  
  strXML = strXML & " </graph>" 
	
'	'Create the chart - Column 3D Chart with data contained in strXML
'		Call renderChart("../../FusionCharts/FCF_Column3D.swf", "", strXML, "productSales", 600, 400)
		Call renderChart("../FusionCharts/FCF_MSColumn3D.swf", "", strXML, "productSales", 800, 350)
	
	set l_rs = Nothing
	cn.Close
	set cn = Nothing	
		

	
	case 3

l_sql = " select estadoInstalacion.descripcionEstadoInsta "
l_sql = l_sql & " ,estadoInstalacion.orden,COUNT(*) as cantidad "
l_sql = l_sql & " from estadoInstalacion "
l_sql = l_sql & " left join detalleVentas on detalleVentas.idestadoInstalacion = estadoInstalacion.id "
l_sql = l_sql & " left join ventas on ventas.id = detalleVentas.idVenta "
l_sql = l_sql & " AND ventas.fecha >=  " & cambiafecha(l_fechadesde,"YMD",true) 
l_sql = l_sql & " AND ventas.fecha <=  " & cambiafecha(l_fechahasta ,"YMD",true) 
l_sql = l_sql & " AND estadoInstalacion.empnro = " & Session("empnro")			
l_sql = l_sql & " group by estadoInstalacion.descripcionEstadoInsta "
l_sql = l_sql & " ,estadoInstalacion.orden " 
l_sql = l_sql & " union  "
l_sql = l_sql & " select 'Vencidas' descripcionEstadoInsta "
l_sql = l_sql & " ,99999 orden " 
l_sql = l_sql & " ,COUNT(*) as cantidad "
l_sql = l_sql & " from detalleVentas "	
l_sql = l_sql & " inner join estadoInstalacion on detalleVentas.idestadoInstalacion = estadoInstalacion.id "
l_sql = l_sql & " inner join ventas on ventas.id = detalleVentas.idVenta "
l_sql = l_sql & " WHERE  (detalleVentas.fechaProgramadaInstalacion is null " 
l_sql = l_sql & " or detalleVentas.fechaProgramadaInstalacion < GETDATE() )"
l_sql = l_sql & " and estadoInstalacion.codigo <> 'F'"
l_sql = l_sql & " AND ventas.fecha >=  " & cambiafecha(l_fechadesde,"YMD",true) 
l_sql = l_sql & " AND ventas.fecha <=  " & cambiafecha(l_fechahasta ,"YMD",true) 
l_sql = l_sql & " AND detalleVentas.empnro = " & Session("empnro")
l_sql = l_sql & " order by orden "

'response.write l_sql&"</br>"
rsOpen l_rs, cn, l_sql, 0 


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

	encabezado3


		
	
		i = 0
    do while not l_rs.eof
	
		'response.write l_rs(1) & "<br>"
		'response.write l_rs(2) & "<br>"
		'response.write l_rs(3) & "<br>"
		'response.end
	
	    arrData(i,1) = l_rs("descripcionEstadoInsta")
		arrData(i,2) = l_rs("cantidad_compra")
		arrData(i,3) = l_rs("cantidad")
	%>
	    <tr>
			
	        <td align="center"><%= l_rs("descripcionEstadoInsta") %></td>	
			<td align="center"><%= l_rs("cantidad") %></td>												   
	    </tr>
		
		 <tr>
		 <td align="center" colspan=15>
	<%
		i = i + 1	
		l_rs.movenext
	loop
	

	'Now, we need to convert this data into XML. We convert using string concatenation.
	Dim strXML3
	Dim strXML_C11
	Dim strXML_1P1
	Dim strXML_R11
	'Initialize <graph> element
	
strXML = " <graph caption='Instalaciones' decimalPrecision='0' showPercentageValues='0' showNames='1' numberPrefix='' showValues='1' showPercentageInLabel='0' pieYScale='45' pieBorderAlpha='100' pieRadius='100' animation='1' shadowXShift='4' shadowYShift='4' shadowAlpha='40' pieFillAlpha='95' pieBorderColor='FFFFFF'> "
  	
'strXML = strXML & " <set name='USA' value='20' isSliced='1'/> "
'strXML = strXML & " <set name='France' value='7'/> "
'strXML = strXML & " </graph> "
	
	
	
	'Convert data to XML and append
	For i=0 to UBound(arrData)-1
'		'add values using <set name='...' value='...' color='...'/>
'		'strXML = strXML & "<set name='" & arrData(i,1) & "' value='" & replace(arrData(i,2),",",".") & "' color='" & getFCColor() & "' />"
'		'strXML = strXML & "<set name='" & arrData(i,1) & "' value='" & replace(arrData(i,3),",",".") & "' color='" & getFCColor() & "' />"
'   		strXML_C11 = strXML_C11 & "   <category name='" & arrData(i,1) & "' hoverText=' "& arrData(i,1) &"'/> "		
'		strXML_1P1 = strXML_1P1 & "    <set value='" & arrData(i,2) & "' /> "
'		strXML_R11 = strXML_R11 & "    <set value='" & arrData(i,3) & "' /> "
		strXML = strXML & " <set name='" & arrData(i,1) &"' value='" & arrData(i,3) &"'/> "
	Next
	'Close <graph> element
	strXML = strXML & "</graph>"

	'response.write strXML & "-"
	'response.end
	
	
	'strXML = "<graph xaxisname='Continent' yaxisname='Export' hovercapbg='DEDEBE' hovercapborder='889E6D' rotateNames='0' animation='1' yAxisMaxValue='100' numdivlines='9' divLineColor='CCCCCC' divLineAlpha='80' decimalPrecision='0' showAlternateVGridColor='1' AlternateVGridAlpha='30' AlternateVGridColor='CCCCCC' caption='Global Export' subcaption='In Millions Tonnes per annum pr Hectare' > "
   'strXML = strXML & "<categories font='Arial' fontSize='11' fontColor='000000'> "
   'strXML = strXML & "   <category name='N. America' hoverText='North America'/> "
   'strXML = strXML & "   <category name='Asia' /> "
   'strXML = strXML & "   <category name='Europe' /> "
   'strXML = strXML & "   <category name='Australia' /> "
   'strXML = strXML & "   <category name='Africa' /> "
'   strXML = strXML & "<categories font='tahoma' fontSize='11' fontColor='000000'> "	
'   strXML = strXML & strXML_C11
'   strXML = strXML & "</categories> "
'  strXML = strXML & " <dataset seriesname='Proyectado' color='FDC12E' alpha='70'> " 
'    strXML = strXML & strXML_1P1
 ' strXML = strXML & " </dataset> "
 ' strXML = strXML & " <dataset seriesname='Real' color='56B9F9' showValues='1' alpha='70'> "
  ''strXML = strXML & strXML_R11
  'strXML = strXML & " </dataset> "
  'strXML = strXML & " </graph>" 
	
'	'Create the chart - Column 3D Chart with data contained in strXML
'		Call renderChart("../../FusionCharts/FCF_Column3D.swf", "", strXML, "productSales", 600, 400)
		Call renderChart("../FusionCharts/FCF_Pie2D.swf", "", strXML, "productSales", 800, 350)
	
	set l_rs = Nothing
	cn.Close
	set cn = Nothing	
	
case 4 'Reporte de caja

	
	l_sql =  "		  select mediosdepago.titulo	as medio_pago "
	l_sql = l_sql & " 		,sum(cajaMovimientos.monto) "
	l_sql = l_sql & " 		,SUM(	case  "
	l_sql = l_sql & " 					when cajaMovimientos.tipo = 'E' then cajaMovimientos.monto  "
	l_sql = l_sql & " 					else 0 "
	l_sql = l_sql & " 				end "
	l_sql = l_sql & " 			)																		as	total_entradas "
	l_sql = l_sql & " 		,SUM(	case  "
	l_sql = l_sql & " 					when cajaMovimientos.tipo = 'S' then cajaMovimientos.monto "
	l_sql = l_sql & " 					else 0 "
	l_sql = l_sql & " 				end "
	l_sql = l_sql & " 			)																		as	total_salidas	 "
	l_sql = l_sql & " 		,SUM(	case  "
	l_sql = l_sql & " 					when cajaMovimientos.tipo = 'S' then -cajaMovimientos.monto "
	l_sql = l_sql & " 					else cajaMovimientos.monto "
	l_sql = l_sql & " 				end "
	l_sql = l_sql & " 			)																		as	saldo						 "
	l_sql = l_sql & " from cajaMovimientos "
	l_sql = l_sql & " inner join mediosdepago on mediosdepago.id = cajaMovimientos.idmedioPago "
	l_sql = l_sql & " WHERE cajaMovimientos.fecha >= "& cambiafecha(l_fechadesde,"YMD",true)
	l_sql = l_sql & " AND  cajaMovimientos.fecha <=  "& cambiafecha(l_fechahasta ,"YMD",true)
	l_sql = l_sql & " AND cajaMovimientos.empnro =  " & Session("empnro") 
	l_sql = l_sql & " group by mediosdepago.titulo "

	'response.write l_sql
	rsOpen l_rs, cn, l_sql, 0 


 %>

<body leftmargin="0" rightmargin="0" topmargin="0" bottommargin="0" onload="//javascript:parent.Buscar();">
<table>
    <tr>
        <td colspan="6">&nbsp;</td>
    </tr>
	<tr>
        <td  colspan="6" align="center" ><h3>Movimientos de Caja desde:&nbsp;<%= l_fechadesde %>&nbsp; al <%= l_fechahasta %>&nbsp;&nbsp;</h3></td>	
    </tr>

<% 	

	encabezado4


		
	
		i = 0
    do while not l_rs.eof	
		
		arrData(i,1) = l_rs("medio_pago")
	    arrData(i,2) = round(l_rs("total_entradas"))
		arrData(i,3) = round(l_rs("total_salidas"))
		arrData(i,4) = round(l_rs("saldo"))
		
	%>
	    <tr>
			
	        <td align="center"><%= l_rs("medio_pago") %></td>
	        <td align="center"><%= l_rs("total_entradas") %></td>	
			<td align="center"><%= l_rs("total_salidas")%></td>
			<td align="center"><%= l_rs("saldo")%></td>
												   
	    </tr>
		
		 <tr>
		 <td align="center" colspan=15>
	<%
		i = i + 1	
		l_rs.movenext
	loop
	

	'Now, we need to convert this data into XML. We convert using string concatenation.
	Dim strXML14
	Dim strXML_C14
	Dim strXML_1P4
	Dim strXML_R14
	Dim strXML_s14
	'Initialize <graph> element
	'eugenio esta linea me paece al pedo porque se reasigna en la isntruccion siguiente strXML = "<graph caption='' subCaption='' yaxisname='Saldo' xaxisname='' formatNumberScale='0' decimalPrecision='2' showNames='1' showValues='1' showPercentageInLabel ='1' showAlternateVGridColor='1' alternateVGridAlpha='10' alternateVGridColor='AFD8F8' numDivLines='4' decimalPrecision='0' canvasBorderThickness='1' canvasBorderColor='114B78' baseFontColor='114B78' hoverCapBorderColor='114B78' hoverCapBgColor='E7EFF6'> "
	
	strXML = "<graph xaxisname='Entradas' yaxisname='Saldo' hovercapbg='DEDEBE' hovercapborder='889E6D' rotateNames='0' animation='1' yAxisMaxValue='0' numdivlines='9' divLineColor='CCCCCC' divLineAlpha='80' decimalPrecision='0' showAlternateVGridColor='1' AlternateVGridAlpha='30' AlternateVGridColor='CCCCCC' caption='Caja' subcaption='' > "

	
	'Convert data to XML and append
	For i=0 to UBound(arrData)-1
		'add values using <set name='...' value='...' color='...'/>
   		strXML_C14 = strXML_C14 & "   <category name='" & arrData(i,1) & "' hoverText=' "& arrData(i,1) &"'/> "		
		strXML_1P4 = strXML_1P4 & "    <set value='" & arrData(i,2) & "' /> "
		strXML_R14 = strXML_R14 & "    <set value='" & arrData(i,3) & "' /> "
		strXML_s14 = strXML_s14 & "    <set value='" & arrData(i,4) & "' /> "
	Next
	'Close <graph> element
	'strXML = strXML & "</graph>"

	'response.write strXML & "-"
	'response.end
	
	
   strXML = strXML & "<categories font='tahoma' fontSize='11' fontColor='000000'> "	
   strXML = strXML & strXML_C14
   strXML = strXML & "</categories> "
  strXML = strXML & " <dataset seriesname='Entradas' color='FDC12E' alpha='70'> " 
    strXML = strXML & strXML_1P4
  strXML = strXML & " </dataset> "
  strXML = strXML & " <dataset seriesname='Salidas' color='8E468E' showValues='1' alpha='70'> "
  strXML = strXML & strXML_R14
  strXML = strXML & " </dataset> "
  strXML = strXML & " <dataset seriesname='Saldo' color='B3AA00' showValues='1' alpha='70'> "
  strXML = strXML & strXML_s14
  strXML = strXML & " </dataset> "  
  strXML = strXML & " </graph>" 
  
	'response.write "hola"&strXML
	'Create the chart - Column 3D Chart with data contained in strXML
		Call renderChart("../FusionCharts/FCF_MSColumn3D.swf", "", strXML, "productSales", 800, 350)
	
	set l_rs = Nothing
	cn.Close
	set cn = Nothing	
	
case 5 'Reporte de caja	por responsable
	
	l_sql =  "		  select mediosdepago.titulo	as medio_pago , responsablesCaja.nombre as nombre_responsable,responsablesCaja.iniciales + ' - '+mediosdepago.titulo as ini_med_pago "
	l_sql = l_sql & " 		,sum(cajaMovimientos.monto) "
	l_sql = l_sql & " 		,SUM(	case  "
	l_sql = l_sql & " 					when cajaMovimientos.tipo = 'E' then cajaMovimientos.monto  "
	l_sql = l_sql & " 					else 0 "
	l_sql = l_sql & " 				end "
	l_sql = l_sql & " 			)																		as	total_entradas "
	l_sql = l_sql & " 		,SUM(	case  "
	l_sql = l_sql & " 					when cajaMovimientos.tipo = 'S' then cajaMovimientos.monto "
	l_sql = l_sql & " 					else 0 "
	l_sql = l_sql & " 				end "
	l_sql = l_sql & " 			)																		as	total_salidas	 "
	l_sql = l_sql & " 		,SUM(	case  "
	l_sql = l_sql & " 					when cajaMovimientos.tipo = 'S' then -cajaMovimientos.monto "
	l_sql = l_sql & " 					else cajaMovimientos.monto "
	l_sql = l_sql & " 				end "
	l_sql = l_sql & " 			)																		as	saldo						 "
	l_sql = l_sql & " from cajaMovimientos "
	l_sql = l_sql & " inner join mediosdepago on mediosdepago.id = cajaMovimientos.idmedioPago "
	l_sql = l_sql & " inner join responsablesCaja on responsablesCaja.id = cajaMovimientos.idresponsable "
	l_sql = l_sql & " WHERE cajaMovimientos.fecha >= "& cambiafecha(l_fechadesde,"YMD",true)
	l_sql = l_sql & " AND  cajaMovimientos.fecha <=  "& cambiafecha(l_fechahasta ,"YMD",true)
	l_sql = l_sql & " AND cajaMovimientos.empnro =  " & Session("empnro") 
	l_sql = l_sql & " group by mediosdepago.titulo , responsablesCaja.nombre ,responsablesCaja.iniciales "

	'response.write l_sql
	rsOpen l_rs, cn, l_sql, 0 


 %>
<body leftmargin="0" rightmargin="0" topmargin="0" bottommargin="0" onload="//javascript:parent.Buscar();">
<table>
    <tr>
        <td colspan="6">&nbsp;</td>
    </tr>
	<tr>
        <td  colspan="6" align="center" ><h3>Movimientos de Caja desde:&nbsp;<%= l_fechadesde %>&nbsp; al <%= l_fechahasta %>&nbsp;&nbsp;</h3></td>	
    </tr>

<% 
	encabezado5


		
	
		i = 0
    do while not l_rs.eof	
		
		arrData(i,1) = l_rs("ini_med_pago")
		'arrData(i,2) = l_rs("medio_pago")
	    arrData(i,3) = round(l_rs("total_entradas"),0)
		arrData(i,4) = round(l_rs("total_salidas"),0)
		arrData(i,5) = round(l_rs("saldo"),0)
%>	
	    <tr>
			
	        <td align="center"><%= l_rs("nombre_responsable") %></td>
			<td align="center"><%= l_rs("medio_pago") %></td>
	        <td align="center"><%= l_rs("total_entradas") %></td>	
			<td align="center"><%= l_rs("total_salidas")%></td>
			<td align="center"><%= l_rs("saldo")%></td>
												   
	    </tr>
		
		 <tr>
		 <td align="center" colspan=15>	
<% 
		i = i + 1	
		l_rs.movenext
	loop

		'Now, we need to convert this data into XML. We convert using string concatenation.
	' Dim strXML14
	' Dim strXML_C14
	' Dim strXML_1P4
	' Dim strXML_R14
	' Dim strXML_s14
	
	'Initialize <graph> element
	'eugenio esta linea me paece al pedo porque se reasigna en la isntruccion siguiente strXML = "<graph caption='' subCaption='' yaxisname='Saldo' xaxisname='' formatNumberScale='0' decimalPrecision='2' showNames='1' showValues='1' showPercentageInLabel ='1' showAlternateVGridColor='1' alternateVGridAlpha='10' alternateVGridColor='AFD8F8' numDivLines='4' decimalPrecision='0' canvasBorderThickness='1' canvasBorderColor='114B78' baseFontColor='114B78' hoverCapBorderColor='114B78' hoverCapBgColor='E7EFF6'> "
	
	strXML = "<graph xaxisname='Entradas' yaxisname='Saldo' hovercapbg='DEDEBE' hovercapborder='889E6D' rotateNames='0' animation='1' yAxisMaxValue='0' numdivlines='9' divLineColor='CCCCCC' divLineAlpha='80' decimalPrecision='0' showAlternateVGridColor='1' AlternateVGridAlpha='30' AlternateVGridColor='CCCCCC' caption='Caja' subcaption='' > "

	
	'Convert data to XML and append
	For i=0 to UBound(arrData)-1
		'add values using <set name='...' value='...' color='...'/>
   		strXML_C14 = strXML_C14 & "   <category name='" & arrData(i,1) & "' hoverText=' "& arrData(i,1) &"'/> "		
		'strXML_1P4 = strXML_1P4 & "    <set value='" & arrData(i,2) & "' /> "
		strXML_1P4 = strXML_1P4 & "    <set value='" & arrData(i,3) & "' /> "
		strXML_R14 = strXML_R14 & "    <set value='" & arrData(i,4) & "' /> "
		strXML_s14 = strXML_s14 & "    <set value='" & arrData(i,5) & "' /> "
	Next
	'Close <graph> element
	'strXML = strXML & "</graph>"

	'response.write strXML & "-"
	'response.end
	
	
   strXML = strXML & "<categories font='tahoma' fontSize='11' fontColor='000000'> "	
   strXML = strXML & strXML_C14
   strXML = strXML & "</categories> "
  strXML = strXML & " <dataset seriesname='Entradas' color='FDC12E' alpha='70'> " 
    strXML = strXML & strXML_1P4
  strXML = strXML & " </dataset> "
  strXML = strXML & " <dataset seriesname='Salidas' color='8E468E' showValues='1' alpha='70'> "
  strXML = strXML & strXML_R14
  strXML = strXML & " </dataset> "
  strXML = strXML & " <dataset seriesname='Saldo' color='B3AA00' showValues='1' alpha='70'> "
  strXML = strXML & strXML_s14
  strXML = strXML & " </dataset> "  
  strXML = strXML & " </graph>" 
  
	'response.write "hola"&strXML
	'Create the chart - Column 3D Chart with data contained in strXML
		Call renderChart("../FusionCharts/FCF_MSColumn3D.swf", "", strXML, "productSales", 800, 350)
	
	set l_rs = Nothing
	cn.Close
	set cn = Nothing	
	
end select
	


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
</body>
</html>

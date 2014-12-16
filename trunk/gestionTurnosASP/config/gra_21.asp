<%@ Language=VBScript %>
<!--#include virtual="/serviciolocal/shared/inc/sec.inc"-->
<!--#include virtual="/serviciolocal/shared/inc/const.inc"-->
<!--#include virtual="/serviciolocal/shared/db/conn_db.inc"-->
<!--#include virtual="/serviciolocal/shared/inc/fecha.inc"-->
<HTML>
<HEAD>
	<TITLE>	Estadisticas </TITLE>
	<%
	
	on error goto 0
	
	'You need to include the following JS file, if you intend to embed the chart using JavaScript.
	'Embedding using JavaScripts avoids the "Click to Activate..." issue in Internet Explorer
	'When you make your own charts, make sure that the path to this JS file is correct. Else, you would get JavaScript errors.
	%>	
	<SCRIPT LANGUAGE="Javascript" SRC="../FusionCharts/FusionCharts.js"></SCRIPT>
	<style type="text/css">
	<!--
	body {
		font-family: Arial, Helvetica, sans-serif;1
		font-size: 12px;
	}
	-->
	</style>
</HEAD>
	<%
	'We've included ../Includes/FusionCharts.asp, which contains functions
	'to help us easily embed the charts.
	%>
<!-- #INCLUDE FILE="../Includes/FusionCharts.asp" -->
	<%
	'We've also included ../Includes/FC_Colors.asp, having a list of colors
	'to apply different colors to the chart's columns. We provide a function for it - getFCColor()
	%>
<!-- #INCLUDE FILE="../Includes/FC_Colors.asp" -->

<BODY>
<CENTER>
<!--
<h5>Buques Atendidos</h5>
-->
<%
	'In this example, we plot a single series chart from data contained
	'in an array. The array will have two columns - first one for data label
	'and the next one for data values.
	
	'Let's store the sales data for 6 products in our array). We also store
	'the name of products. 
	Dim arrData()
	Dim l_CantAgencias
	
	Dim l_rs
	Dim l_sql
	Dim l_fecini
	Dim l_fecfin

	Set l_rs = Server.CreateObject("ADODB.RecordSet")
	
	l_fecini = request.querystring("fecini")
	l_fecfin = request.querystring("fecfin")

	l_sql = " SELECT count(distinct(ser_legajo.legpar1)) "
	l_sql = l_sql & " FROM ser_legajo "
	l_sql = l_sql & " INNER JOIN ser_servicio ON ser_legajo.legpar1 = ser_servicio.sercod "
	l_sql = l_sql & " WHERE ser_legajo.legfecing >= " & cambiafecha(l_fecini,"YMD",true)
	l_sql = l_sql & " AND ser_legajo.legfecing <= " & cambiafecha(l_fecfin,"YMD",true)

	l_sql = l_sql & " group by ser_legajo.legpar1 "
	rsOpen l_rs, cn, l_sql, 0
	
	l_CantAgencias = 0
	do while not l_rs.eof
		l_CantAgencias = l_CantAgencias + 1
		l_rs.movenext	
	loop
	l_rs.close

	ReDim arrData(l_CantAgencias,2)
	
	l_sql = " SELECT distinct(ser_legajo.legpar1), ser_servicio.serdes, count(*) "
	l_sql = l_sql & " FROM ser_legajo "
	l_sql = l_sql & " INNER JOIN ser_servicio ON ser_legajo.legpar1 = ser_servicio.sercod "
	l_sql = l_sql & " WHERE ser_legajo.legfecing >= " & cambiafecha(l_fecini,"YMD",true)
	l_sql = l_sql & " AND ser_legajo.legfecing <= " & cambiafecha(l_fecfin,"YMD",true)

	l_sql = l_sql & " group by ser_legajo.legpar1 , ser_servicio.serdes"

	rsOpen l_rs, cn, l_sql, 0
	
	l_fila = 0
	do while not l_rs.eof
		arrData(l_fila,1) = l_rs(1)
		arrData(l_fila,2) = l_rs(2)
		l_fila = l_fila + 1
		l_rs.movenext
	loop

%>
<!--
<graph caption='Fruit Production for March' subCaption='(in Millions)' yaxisname='Fruit' xaxisname='Quantity' showAlternateVGridColor='1' alternateVGridAlpha='10' alternateVGridColor='AFD8F8' numDivLines='4' decimalPrecision='0' canvasBorderThickness='1' canvasBorderColor='114B78' baseFontColor='114B78' hoverCapBorderColor='114B78' hoverCapBgColor='E7EFF6'>
   <set name='Orange' value='23' color='AFD8F8' alpha='70'/> 
   <set name='Apple' value='12' color='F6BD0F' alpha='70'/> 
   <set name='Banana' value='17' color='8BBA00' alpha='70'/> 
   <set name='Mango' value='14' color='A66EDD' alpha='70'/> 
   <set name='Litchi' value='12' color='F984A1' alpha='70'/>
</graph>
-->
<%
'response.end

	'Now, we need to convert this data into XML. We convert using string concatenation.
	Dim strXML, i
	'Initialize <graph> element
	'strXML = "<graph caption='Estadísticas de Pedidos' numberPrefix='$' formatNumberScale='0' decimalPrecision='0'>"
	strXML = "<graph caption='' subCaption='' yaxisname='Cantidad' xaxisname='Servicios' showAlternateVGridColor='1' alternateVGridAlpha='10' alternateVGridColor='AFD8F8' numDivLines='4' decimalPrecision='0' canvasBorderThickness='1' canvasBorderColor='114B78' baseFontColor='114B78' hoverCapBorderColor='114B78' hoverCapBgColor='E7EFF6'> "
	
	'Convert data to XML and append
	For i=0 to UBound(arrData)-1
		'add values using <set name='...' value='...' color='...'/>
		strXML = strXML & "<set name='" & arrData(i,1) & "' value='" & arrData(i,2) & "' color='" & getFCColor() & "' />"
	Next
	'Close <graph> element
	strXML = strXML & "</graph>"

	
'	'Create the chart - Column 3D Chart with data contained in strXML
'		Call renderChart("../../FusionCharts/FCF_Column3D.swf", "", strXML, "productSales", 600, 400)
		Call renderChart("../FusionCharts/FCF_Bar2D.swf", "", strXML, "productSales", 600, 250)
	'renderChartHTML
	set l_rs = Nothing
	cn.Close
	set cn = Nothing	
%>
<BR><BR>
</CENTER>
</BODY>
</HTML>
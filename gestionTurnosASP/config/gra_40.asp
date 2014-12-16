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
	
	Dim l_rs
	Dim l_sql

	ReDim arrData(3,2)
	
	arrData(0,1) = left(request.querystring("nom1"),7)
	arrData(0,2) = request.querystring("val1")
	
	arrData(1,1) = left(request.querystring("nom2"),7)
	arrData(1,2) = request.querystring("val2")
	
	arrData(2,1) = left(request.querystring("nom3"),7)
	arrData(2,2) = request.querystring("val3")

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
	strXML = "<graph caption='' subCaption='' yaxisname='' formatNumberScale='0' formatNumber='0 '  xaxisname='' showAlternateVGridColor='1' alternateVGridAlpha='10' alternateVGridColor='AFD8F8' numDivLines='4' decimalPrecision='0' canvasBorderThickness='1' canvasBorderColor='114B78' baseFontColor='114B78' hoverCapBorderColor='114B78' hoverCapBgColor='E7EFF6'> "
	
	'Convert data to XML and append
	For i=0 to UBound(arrData)-1
		'add values using <set name='...' value='...' color='...'/>
		strXML = strXML & "<set name='" & arrData(i,1) & "' value='" & arrData(i,2) & "' color='" & getFCColor() & "' />"
	Next
	'Close <graph> element
	strXML = strXML & "</graph>"

	
'	'Create the chart - Column 3D Chart with data contained in strXML
'		Call renderChart("../../FusionCharts/FCF_Column3D.swf", "", strXML, "productSales", 600, 400)
		Call renderChart("../FusionCharts/FCF_Column2D.swf", "", strXML, "productSales", 350, 150)
	'renderChartHTML
	set l_rs = Nothing
	cn.Close
	set cn = Nothing
%>
<BR><BR>
</CENTER>
</BODY>
</HTML>
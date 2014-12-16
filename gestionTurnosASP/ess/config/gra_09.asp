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
<%
	Dim arrData(12,2)
	
	Dim l_rs
	Dim l_sql
	
	Dim l_cadena
	Dim l_lista
	Dim l_elem
	Dim i
	
	l_cadena = request.querystring("cadena")
	
	l_lista = Split(l_cadena,",")
	
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
	strXML = "<graph caption='' subCaption='' yaxisname='' xaxisname='' formatNumberScale='0' decimalSeparator=',' thousandSeparator='.'   showAlternateVGridColor='1' alternateVGridAlpha='10' alternateVGridColor='AFD8F8' numDivLines='4' decimalPrecision='0' canvasBorderThickness='1' canvasBorderColor='114B78' baseFontColor='114B78' hoverCapBorderColor='114B78' hoverCapBgColor='E7EFF6'> "
	
	'Convert data to XML and append
	For i=0 to UBound(arrData)-1
		'add values using <set name='...' value='...' color='...'/>
		strXML = strXML & "<set name='" & arrData(i,1) & "' value='" & arrData(i,2) & "' color='" & getFCColor() & "' />"
	Next
	'Close <graph> element
	strXML = strXML & "</graph>"

	'response.write strXML & "-"
	'response.end
	
	
'	'Create the chart - Column 3D Chart with data contained in strXML
'		Call renderChart("../../FusionCharts/FCF_Column3D.swf", "", strXML, "productSales", 600, 400)
		Call renderChart("../FusionCharts/FCF_Bar2D.swf", "", strXML, "productSales", 500, 300)
	
	set l_rs = Nothing
	cn.Close
	set cn = Nothing	
%>
<BR><BR>
</CENTER>
</BODY>
</HTML>
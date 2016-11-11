<% Option Explicit %>
<!--#include virtual="/trivoliSwimming/shared/db/conn_db.inc"-->
<% 
'Archivo: bancos_con_01.asp
'Descripción: Grilla Administración de bancos
'Autor : Trivoli
'Fecha: 31/05/2015

Dim l_rs
Dim l_sql
Dim l_filtro
Dim l_orden
Dim l_sqlfiltro
Dim l_sqlorden
Dim l_totvol
Dim l_cant

Dim l_primero

l_filtro = request("filtro")
l_orden  = request("orden")

if l_orden = "" then
  l_orden = " ORDER BY bancos.nombre_banco "
end if
%>

<!DOCTYPE HTML PUBLIC "-//IETF//DTD HTML//EN">
<html>
<head>
    <title id='Description'>Compras Franquicias</title>
    <link rel="stylesheet" href="../../js/jqwidgets/styles/jqx.base.css" type="text/css" />
    
	<!--INICIO MODIFICACION TRIVOLI-->
	<!--<script type="text/javascript" src="../../scripts/jquery-1.11.1.min.js"></script>-->	
	<link rel="stylesheet" href="../../js/themes/smoothness/jquery-ui.css" />
	<script src="../../js/jquery.min.js"></script>
	<script src="../../js/jquery-ui.js"></script>
	<!--FIN MODIFICACION TRIVOLI-->
	
    <script type="text/javascript" src="../../js/jqwidgets/jqxcore.js"></script>
    <script type="text/javascript" src="../../js/jqwidgets/jqxbuttons.js"></script>
    <script type="text/javascript" src="../../js/jqwidgets/jqxscrollbar.js"></script>
    <script type="text/javascript" src="../../js/jqwidgets/jqxmenu.js"></script>
    <script type="text/javascript" src="../../js/jqwidgets/jqxgrid.js"></script>
    <script type="text/javascript" src="../../js/jqwidgets/jqxgrid.selection.js"></script>
    <script type="text/javascript" src="../../js/jqwidgets/jqxgrid.columnsresize.js"></script>
    <script type="text/javascript" src="../../js/jqwidgets/jqxgrid.pager.js"></script>
    <script type="text/javascript" src="../../js/jqwidgets/jqxlistbox.js"></script>
    <script type="text/javascript" src="../../js/jqwidgets/jqxdropdownlist.js"></script>
    <script type="text/javascript" src="../../js/jqwidgets/jqxdata.js"></script>    
	<script type="text/javascript" src="compras_rep.js"></script>

</head>
<body class='default'>
    <div id='jqxWidget' style="font-size: 13px; font-family: Verdana; float: left;">
        <h3>Compras</h3>
        <div id="comprasGrid">
        </div>
		<table width="100%">
		    <tr>
				<td>
					<h3>Detalle Compra</h3>
					<div id="detalleComprasGrid">
					</div>				
				</td>
				<td>
					<h3>Detalle Pagos</h3>
					<div id="detallePagosGrid">
					</div>					
				</td>		
			</tr>
		</table>

    </div>
</body>
</html>
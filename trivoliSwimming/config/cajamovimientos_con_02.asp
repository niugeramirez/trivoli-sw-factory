
<% Option Explicit %>
<!--#include virtual="/trivoliSwimming/shared/db/conn_db.inc"-->
<% 

on error goto 0

'Datos del formulario

dim l_id
dim l_fecha
dim l_tipoes
dim l_idtipomovimiento 
dim l_detalle
dim l_idunidadnegocio
dim l_idmediopago
dim l_idcheque
dim l_monto
dim l_idresponsable
dim l_idcompraorigen
dim l_idventaorigen
Dim l_mediodepagocheque

Dim l_compraorigen
Dim l_ventaorigen
Dim l_cheque_nom

Dim p_flagcompra
Dim p_flagventa

Dim l_p_id_venta
Dim l_p_id_compra

'ADO
Dim l_tipo
Dim l_sql
Dim l_rs

l_tipo = request.querystring("tipo")
l_p_id_venta = request.querystring("p_id_venta")
l_p_id_compra = request.querystring("p_id_compra")
'response.write  "p_id_compra "&l_p_id_compra&"</br>"

Set l_rs = Server.CreateObject("ADODB.RecordSet")

'obtengo el Medio de Pago Obra Social
l_sql = "SELECT * "
l_sql = l_sql & " FROM mediosdepago "
l_sql  = l_sql  & " WHERE flag_cheque = -1 " 
l_sql = l_sql & " AND empnro = " & Session("empnro")
rsOpen l_rs, cn, l_sql, 0 
l_mediodepagocheque = 0
if not l_rs.eof then
	l_mediodepagocheque = l_rs("id")	
end if
l_rs.Close

%>
<html>
<head>
<meta http-equiv="Content-Type" content="text/html; charset=iso-8859-1">

<style type="text/css">
#contenedor {
    display: table;
	
}
#contenidos {
    display: table-row;
}
#columna1 {
    display: table-cell;
	COLOR: black;
	FONT-FAMILY: Verdana;
	FONT-SIZE: 08pt;
	/*BACKGROUND-COLOR: #E4FEF9;*/
	padding : 2;
	padding-left : 5;
	width: 110px;
}
#columna2, #columna3 {
    display: table-cell;
	COLOR: black;
	FONT-FAMILY: Verdana;
	FONT-SIZE: 08pt;
	/*BACKGROUND-COLOR: #E4FEF9;*/
	padding : 2;
	padding-left : 5;
	width: 500px;
}
</style>


<!-- Comienzo Datepicker -->
<script>
$(function () {
/*$.datepicker.setDefaults($.datepicker.regional["es"]);
$("#datepicker").datepicker({
firstDay: 1
});*/

		
$( "#fecha" ).datepicker({
	showOn: "button",
	buttonImage: "../shared/images/calendar16.png",
	buttonImageOnly: true
});

});
</script>
<!-- Final Datepicker -->

<script>
function ctrolmediodepago(){

	if (document.datos_02_mc.mediodepagocheque.value == document.datos_02_mc.idmediopago.value) {			
			//document.datos_02_mc.idcheque.disabled = false;	
			mostrar('cheque');						
		}
		else {
		
			//document.datos_02_mc.idcheque.disabled = true;	
			cerrar('cheque');								
	
		}	

}

function ctrolcheque(){

document.valida.location = "importecheque_con_00.asp?id=" + document.datos_02_mc.idcheque.value ;	

}

function actualizarimporte(p_importe){	
	document.datos_02_mc.monto.value = p_importe;
}


function ctroltipomovimiento(){

	document.valida.location = "flagtipomovimiento_con_00.asp?id=" + document.datos_02_mc.idtipomovimiento.value ;	

}

function actualizarflag (p_flagcompra, p_flagventa){	
	
		if (p_flagcompra == -1 ) {			
			document.datos_02_mc.idcompraorigen.disabled = false;		
			mostrar('compraorigen');					
		}
		else {			
			document.datos_02_mc.idcompraorigen.disabled = true;	
			document.datos_02_mc.idcompraorigen.value = 0;
			cerrar('compraorigen');					
		};

		if (p_flagventa == -1 ) {			
			document.datos_02_mc.idventaorigen.disabled = false;	
			mostrar('ventaorigen');						
		}
		else {			
			document.datos_02_mc.idventaorigen.disabled = true;	
			document.datos_02_mc.idventaorigen.value = 0;
			cerrar('ventaorigen');	
			
		};	
	
}

</script>

<script languague="javascript">
        function mostrar(nombrediv) {
			
            div = document.getElementById(nombrediv);
            div.style.display = '';
        }

        function cerrar(nombrediv) {
		
            div = document.getElementById(nombrediv);
            div.style.display = 'none';
        }
</script>

</head>

<% 
select Case l_tipo
	Case "A":
 	    	l_fecha    	   = date()
			if l_p_id_compra <> "" then 		
				l_tipoes    = "S"
			else
				l_tipoes    = "E"
			end if
			
			l_idtipomovimiento       = "0"
			'Si viene el parametro de venta entonces ya selecciono el tipo de movimeinto de venta venta
			if l_p_id_venta <> "" then 						
				l_idtipomovimiento = "1"
				
				Set l_rs = Server.CreateObject("ADODB.RecordSet")
				l_sql = "SELECT  * "
				l_sql = l_sql & " FROM tiposMovimientoCaja  "								
				l_sql = l_sql  & " WHERE tiposMovimientoCaja.flagVenta = -1 "
				l_sql = l_sql & " and tiposMovimientoCaja.empnro = " & Session("empnro") 
				rsOpen l_rs, cn, l_sql, 0 
				if not l_rs.eof then
					l_idtipomovimiento           = l_rs("id") 		
				end if
				l_rs.Close
	
			else
				if l_p_id_compra <> "" then 						
					l_idtipomovimiento = "1"
					
					Set l_rs = Server.CreateObject("ADODB.RecordSet")
					l_sql = "SELECT  * "
					l_sql = l_sql & " FROM tiposMovimientoCaja  "								
					l_sql = l_sql  & " WHERE tiposMovimientoCaja.flagCompra = -1 "
					l_sql = l_sql & " and tiposMovimientoCaja.empnro = " & Session("empnro") 
					rsOpen l_rs, cn, l_sql, 0 
					if not l_rs.eof then
						l_idtipomovimiento           = l_rs("id") 		
					end if
					l_rs.Close
		
				else			
					l_idtipomovimiento = "0"
				end if
			end if
			
			l_detalle	     = ""
			l_idunidadnegocio    = "0"
			l_idmediopago 		 = "0"
			l_cheque_nom         = ""
	    	l_idcheque = "0"
	    	l_monto  = "0"
			l_idresponsable = "0"
			
			'Si viene el parametro de compra entonces ya selecciono la venta
			if l_p_id_compra <> "" then 						
				l_idcompraorigen = l_p_id_compra
			else
				l_idcompraorigen = "0"
			end if						
			
			'Si viene el parametro de venta entonces ya selecciono la venta
			if l_p_id_venta <> "" then 						
				l_idventaorigen = l_p_id_venta
			else
				l_idventaorigen = "0"
			end if
			
			
	Case "M":
		Set l_rs = Server.CreateObject("ADODB.RecordSet")
		l_id = request.querystring("cabnro")
		l_sql = "SELECT  cajamovimientos.* , proveedores.nombre prov_nom , compras.fecha comp_fec, clientes.nombre cli_nom , ventas.fecha venta_fec, cheques.numero cheque_num , cheques.fecha_emision ,  bancos.nombre_banco "
		l_sql = l_sql & " FROM cajamovimientos  "
		l_sql = l_sql & " LEFT JOIN compras ON compras.id = cajamovimientos.idcompraorigen  "		
		l_sql = l_sql & " LEFT JOIN proveedores ON proveedores.id = compras.idproveedor "				
		l_sql = l_sql & " LEFT JOIN ventas ON ventas.id = cajamovimientos.idventaorigen  "				
		l_sql = l_sql & " LEFT JOIN clientes ON clientes.id = ventas.idcliente  "						
		l_sql = l_sql & " LEFT JOIN cheques ON cheques.id = cajamovimientos.idcheque  "	
		l_sql = l_sql & " LEFT JOIN bancos ON bancos.id = cheques.id_banco "
		l_sql  = l_sql  & " WHERE cajamovimientos.id = " & l_id
		rsOpen l_rs, cn, l_sql, 0 
		if not l_rs.eof then
 	    	l_fecha      		     = l_rs("fecha")
			l_tipoes	     			 = l_rs("tipo")
			l_idtipomovimiento		 = l_rs("idtipomovimiento")
			l_detalle   			 = l_rs("detalle")
			l_idunidadnegocio  		 = l_rs("idunidadnegocio")
			l_idmediopago            = l_rs("idmediopago")
			l_cheque_nom             = l_rs("cheque_num") & " - " & l_rs("nombre_banco")  & " - " & l_rs("fecha_emision") 
			l_idcheque				 = l_rs("idcheque")
	    	l_monto 				 = l_rs("monto")
	    	l_idresponsable  		 = l_rs("idresponsable")
			l_idcompraorigen         = l_rs("idcompraorigen")				
			l_idventaorigen		     = l_rs("idventaorigen")
			
			
		end if
		l_rs.Close
end select

'Inicializaciones generales mas alla de si es Alta o modificacion de registro
if l_idventaorigen <> "0" then
	Set l_rs = Server.CreateObject("ADODB.RecordSet")
	l_sql = "SELECT  clientes.nombre cli_nom , ventas.fecha venta_fec"
	l_sql = l_sql & " FROM ventas  "			
	l_sql = l_sql & " LEFT JOIN clientes ON clientes.id = ventas.idcliente  "						
	l_sql  = l_sql  & " WHERE ventas.id = " & l_idventaorigen
	rsOpen l_rs, cn, l_sql, 0 
	if not l_rs.eof then
		l_ventaorigen           = l_rs("cli_nom") & " - " & l_rs("venta_fec")			
	end if
	l_rs.Close
end if

if l_idcompraorigen <> "0" then
	Set l_rs = Server.CreateObject("ADODB.RecordSet")
	l_sql = "SELECT  proveedores.nombre prov_nom , compras.fecha comp_fec "
	l_sql = l_sql & " FROM compras "			
	l_sql = l_sql & " LEFT JOIN proveedores ON proveedores.id = compras.idproveedor "						
	l_sql  = l_sql  & " WHERE compras.id  = " & l_idcompraorigen
	rsOpen l_rs, cn, l_sql, 0 
	if not l_rs.eof then
		l_compraorigen           = l_rs("prov_nom") & " - " & l_rs("comp_fec")				
	end if
	l_rs.Close
end if
'Fin inicializacion generales 
%>
<body leftmargin="0" rightmargin="0" topmargin="0" bottommargin="0" onload="javascript:document.datos_02_mc.numero.focus();">	
	<form name="datos_02_mc" id="datos_02_mc" action = "Javascript:Submit_Formulario_mc();" onkeypress="if (event.keyCode == 13) {event.preventDefault();Submit_Formulario_mc();}"  target="valida">
		<input type="Hidden" name="id" value="<%= l_id %>">
		<input type="Hidden" name="tipo" value="<%= l_tipo %>">
		<input type="Hidden" name="mediodepagocheque" value="<%= l_mediodepagocheque %>">
		
<div id="contenedor">
    <div id="flotand">
        <div id="columna1" align="right">Fecha:</div>
        <div id="columna2"><input type="text" id="fecha" name="fecha" size="10" maxlength="10" value="<%= l_fecha %>">		</div>
       
    </div>	
	
    <div id="tipo">
        <div id="columna1" align="right">Tipo:</div>
        <div id="columna2">
			<select  name="tipoes" id="tipoes" size="1" style="width:250;">		
				<option value= "E" >Entrada</option>
				<option value= "S" >Salida</option>
			</select>
			<script>document.datos_02_mc.tipoes.value= "<%= l_tipoes%>"</script>
			<%if l_p_id_venta <> "" or l_p_id_compra <> "" then%>
				<script>$('#tipoes option:not(:selected)').attr('disabled',true);</script>
			<%end if%>
		</div>
       
    </div>

    <div id="tipomovimiento">
        <div id="columna1" align="right">Tipo Movimiento:</div>
        <div id="columna2">
			<select name="idtipomovimiento" id="idtipomovimiento" size="1" style="width:250;" onchange="ctroltipomovimiento();">
				<option value="0" selected>&nbsp;Seleccione un Tipo de Movimiento</option>
				<%Set l_rs = Server.CreateObject("ADODB.RecordSet")
				l_sql = "SELECT  * "
				l_sql  = l_sql  & " FROM tiposMovimientoCaja "
				l_sql = l_sql & " where tiposMovimientoCaja.empnro = " & Session("empnro")   
				
				l_sql  = l_sql  & " ORDER BY descripcion "
				rsOpen l_rs, cn, l_sql, 0
				do until l_rs.eof		%>	
				<option value= <%= l_rs("id") %> > 
				<%= l_rs("descripcion") %>  </option>
				<%	l_rs.Movenext
				loop
				l_rs.Close %>
			</select>
			<script>document.datos_02_mc.idtipomovimiento.value= "<%= l_idtipomovimiento%>"</script>	
			<%if l_p_id_venta <> "" or l_p_id_compra <> "" then%>
				<script>$('#idtipomovimiento option:not(:selected)').attr('disabled',true);</script>
			<%end if%>			
		</div>
       
    </div>	
    <div id="compraorigen">
        <div id="columna1" align="right">Compra Origen:</div>
        <div id="columna2">
		<input class="deshabinp" readonly="" type="text" name="compraorigen" id="compraorigen" size="50" maxlength="50" value="<%=l_compraorigen %>">		
									<input type="hidden" name="idcompraorigen" id="idcompraorigen" size="10" maxlength="10" value="<%=l_idcompraorigen %>">		
									<a href="Javascript:BuscarCompraOrigen();"><img src="../shared/images/Buscar_16.png" border="0" title="Buscar Compra Origen"></a>	</div>
       
    </div>		
    <div id="ventaorigen">
        <div id="columna1" align="right">Venta Origen:</div>
        <div id="columna2">
		<input class="deshabinp" readonly="" type="text" name="ventaorigen" id="ventaorigen" size="50" maxlength="50" value="<%=l_ventaorigen %>">		
									<input type="hidden" name="idventaorigen" id="idventaorigen" size="10" maxlength="10" value="<%=l_idventaorigen %>">		
									<a href="Javascript:BuscarVentaOrigen();"><img src="../shared/images/Buscar_16.png" border="0" title="Buscar Venta Origen"></a>	</div>
       
    </div>		
	
    <div id="unidaddenegocio">
        <div id="columna1" align="right">Unidad de Negocio:</div>
        <div id="columna2">
		<select name="idunidadnegocio" size="1" style="width:250;">
										<option value="0" selected>&nbsp;Seleccione una Unidad de Negocio</option>
										<%Set l_rs = Server.CreateObject("ADODB.RecordSet")
										l_sql = "SELECT  * "
										l_sql  = l_sql  & " FROM unidadesNegocio "
										l_sql = l_sql & " where unidadesNegocio.empnro = " & Session("empnro")   
										
										l_sql  = l_sql  & " ORDER BY descripcion "
										rsOpen l_rs, cn, l_sql, 0
										do until l_rs.eof		%>	
										<option value= <%= l_rs("id") %> > 
										<%= l_rs("descripcion") %>  </option>
										<%	l_rs.Movenext
										loop
										l_rs.Close %>
									</select>
									<script>document.datos_02_mc.idunidadnegocio.value= "<%= l_idunidadnegocio%>"</script>		</div>
       
    </div>		
    <div id="mediodepago">
        <div id="columna1" align="right">Medio de Pago:</div>
        <div id="columna2">
		<select name="idmediopago" size="1" style="width:250;" onchange="ctrolmediodepago();">
										<option value="0" selected>&nbsp;Seleccione un Medio de Pago</option>
										<%Set l_rs = Server.CreateObject("ADODB.RecordSet")
										l_sql = "SELECT  * "
										l_sql  = l_sql  & " FROM mediosdepago "
										l_sql = l_sql & " where mediosdepago.empnro = " & Session("empnro")   
										
										l_sql  = l_sql  & " ORDER BY titulo "
										rsOpen l_rs, cn, l_sql, 0
										do until l_rs.eof		%>	
										<option value=<%= l_rs("id") %> > 
										<%= l_rs("titulo") %>  </option>
										<%	l_rs.Movenext
										loop
										l_rs.Close %>
									</select>
									<script>document.datos_02_mc.idmediopago.value= "<%= l_idmediopago %>"</script>	</div>
       
    </div>		
    
    <div id="cheque">
        <div id="columna1" align="right">Cheque:</div>
        <div id="columna2">
			<input class="deshabinp" readonly="" type="text" name="cheque_nom" id="cheque_nom" size="50" maxlength="50" value="<%=l_cheque_nom %>">		
			<input type="hidden" name="idcheque" id="idcheque" size="10" maxlength="10" value="<%=l_idcheque %>">		
			<a href="Javascript:BuscarCheque();"><img src="../shared/images/Buscar_16.png" border="0" title="Buscar Cheque"></a>	
			<a href="Javascript:Editar_Cheque();"><img src="../shared/images/Modificar_16.png" border="0" title="Editar Cheque"></a>			
		</div>
    </div>			
    <div id="monto">
        <div id="columna1" align="right">Monto:</div>
        <div id="columna2">
		<input type="text" name="monto" size="50" maxlength="50" value="<%= l_monto%>">		
									<input type="hidden" name="monto2" value="">	</div>
       
    </div>	
    <div id="responsable">
        <div id="columna1" align="right">Responsable:</div>
        <div id="columna2">
		<select name="idresponsable" size="1" style="width:250;">
										<option value="0" selected>&nbsp;Seleccione un Responsable</option>
										<%Set l_rs = Server.CreateObject("ADODB.RecordSet")
										l_sql = "SELECT  * "
										l_sql  = l_sql  & " FROM responsablesCaja "
										l_sql = l_sql & " where responsablesCaja.empnro = " & Session("empnro")   
										
										l_sql  = l_sql  & " ORDER BY nombre "
										rsOpen l_rs, cn, l_sql, 0
										do until l_rs.eof		%>	
										<option value= <%= l_rs("id") %> > 
										<%= l_rs("nombre") %> </option>
										<%	l_rs.Movenext
										loop
										l_rs.Close %>
									</select>
									<script>document.datos_02_mc.idresponsable.value= "<%= l_idresponsable %>"</script>	</div>
       
    </div>		
    <div id="detalle">
        <div id="columna1" align="right">Detalle:</div>
        <div id="columna2">
		<input type="text" name="detalle" size="70" maxlength="200" value="<%= l_detalle %>">		</div>
       
    </div>		

   
</div>				
		<iframe name="valida"  style="visibility=hidden;" src="" width="0%" height="0%"></iframe> 		
	</form>
<%
set l_rs = nothing
cn.Close
set cn = nothing
%>
<script>
ctrolmediodepago();
ctroltipomovimiento();
</script>
</body>
</html>


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

Dim p_flagcompra
Dim p_flagventa

'ADO
Dim l_tipo
Dim l_sql
Dim l_rs

l_tipo = request.querystring("tipo")

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

	if (document.datos_02.mediodepagocheque.value == document.datos_02.idmediopago.value) {			
			document.datos_02.idcheque.disabled = false;							
		}
		else {
		
			document.datos_02.idcheque.disabled = true;							
	
		}	

}

function ctrolcheque(){

document.valida.location = "importecheque_con_00.asp?id=" + document.datos_02.idcheque.value ;	

}

function actualizarimporte(p_importe){	
	document.datos_02.monto.value = p_importe;
}


function ctroltipomovimiento(){

	document.valida.location = "flagtipomovimiento_con_00.asp?id=" + document.datos_02.idtipomovimiento.value ;	

}

function actualizarflag (p_flagcompra, p_flagventa){	
	
		if (p_flagcompra == -1 ) {			
			document.datos_02.idcompraorigen.disabled = false;		
			mostrar('compraorigen');					
		}
		else {			
			document.datos_02.idcompraorigen.disabled = true;	
			document.datos_02.idcompraorigen.value = 0;
			cerrar('compraorigen');					
		};

		if (p_flagventa == -1 ) {			
			document.datos_02.idventaorigen.disabled = false;	
			mostrar('ventaorigen');						
		}
		else {			
			document.datos_02.idventaorigen.disabled = true;	
			document.datos_02.idventaorigen.value = 0;
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
			l_tipoes    = "E"
			'l_fecha_vencimiento     = ""
			l_idtipomovimiento       = "0"
			l_detalle	     = ""
			l_idunidadnegocio    = "0"
			l_idmediopago 		 = "0"
	    	l_idcheque = "0"
	    	l_monto  = "0"
			l_idresponsable = "0"
			l_idcompraorigen = "0"
			l_idventaorigen = "0"

	Case "M":
		Set l_rs = Server.CreateObject("ADODB.RecordSet")
		l_id = request.querystring("cabnro")
		l_sql = "SELECT  cajamovimientos.* , proveedores.nombre prov_nom , compras.fecha comp_fec, clientes.nombre cli_nom , ventas.fecha venta_fec"
		l_sql = l_sql & " FROM cajamovimientos  "
		l_sql = l_sql & " LEFT JOIN compras ON compras.id = cajamovimientos.idcompraorigen  "		
		l_sql = l_sql & " LEFT JOIN proveedores ON proveedores.id = compras.idproveedor "				
		l_sql = l_sql & " LEFT JOIN ventas ON ventas.id = cajamovimientos.idventaorigen  "				
		l_sql = l_sql & " LEFT JOIN clientes ON clientes.id = ventas.idcliente  "						
		l_sql  = l_sql  & " WHERE cajamovimientos.id = " & l_id
		rsOpen l_rs, cn, l_sql, 0 
		if not l_rs.eof then
 	    	l_fecha      		     = l_rs("fecha")
			l_tipoes	     			 = l_rs("tipo")
			l_idtipomovimiento		 = l_rs("idtipomovimiento")
			l_detalle   			 = l_rs("detalle")
			l_idunidadnegocio  		 = l_rs("idunidadnegocio")
			l_idmediopago            = l_rs("idmediopago")
			l_idcheque				 = l_rs("idcheque")
	    	l_monto 				 = l_rs("monto")
	    	l_idresponsable  		 = l_rs("idresponsable")
			l_idcompraorigen         = l_rs("idcompraorigen")
			l_compraorigen           = l_rs("prov_nom") & " - " & l_rs("comp_fec")
			l_ventaorigen           = l_rs("cli_nom") & " - " & l_rs("venta_fec")			
			l_idventaorigen		     = l_rs("idventaorigen")
			
			
		end if
		l_rs.Close
end select

%>
<body leftmargin="0" rightmargin="0" topmargin="0" bottommargin="0" onload="javascript:document.datos_02.numero.focus();">	
	<form name="datos_02" id="datos_02" action = "Javascript:Submit_Formulario();" onkeypress="if (event.keyCode == 13) {event.preventDefault();Submit_Formulario();}"  target="valida">
		<input type="Hidden" name="id" value="<%= l_id %>">
		<input type="Hidden" name="tipo" value="<%= l_tipo %>">
		<input type="Hidden" name="mediodepagocheque" value="<%= l_mediodepagocheque %>">
		
<div id="contenedor">

    <div id="tipo">
        <div id="columna1" align="right">Tipo:</div>
        <div id="columna2">
		<select name="tipoes" size="1" style="width:250;">		
										<option value= "E" >Entrada</option>
										<option value= "S" >Salida</option>
									</select>
									<script>document.datos_02.tipoes.value= "<%= l_tipoes%>"</script>		</div>
       
    </div>
    <div id="flotand">
        <div id="columna1" align="right">Fecha:</div>
        <div id="columna2"><input type="text" id="fecha" name="fecha" size="10" maxlength="10" value="<%= l_fecha %>">		</div>
       
    </div>	
    <div id="tipomovimiento">
        <div id="columna1" align="right">Tipo Movimiento:</div>
        <div id="columna2">
		<select name="idtipomovimiento" size="1" style="width:250;" onchange="ctroltipomovimiento();">
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
									<script>document.datos_02.idtipomovimiento.value= "<%= l_idtipomovimiento%>"</script>		</div>
       
    </div>	
    <div id="detalle">
        <div id="columna1" align="right">Detalle:</div>
        <div id="columna2">
		<input type="text" name="detalle" size="70" maxlength="200" value="<%= l_detalle %>">		</div>
       
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
									<script>document.datos_02.idunidadnegocio.value= "<%= l_idunidadnegocio%>"</script>		</div>
       
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
									<script>document.datos_02.idmediopago.value= "<%= l_idmediopago %>"</script>	</div>
       
    </div>		
    <div id="cheque">
        <div id="columna1" align="right">Cheque:</div>
        <div id="columna2">
		<select name="idcheque" size="1" style="width:250;" onchange="ctrolcheque();">
										<option value="0" selected>&nbsp;Seleccione un Cheque</option>
										<%Set l_rs = Server.CreateObject("ADODB.RecordSet")
										l_sql = "SELECT  cheques.id, cheques.numero, bancos.nombre_banco , cheques.importe"
										l_sql  = l_sql  & " FROM cheques "
										l_sql  = l_sql  & " LEFT JOIN bancos ON bancos.id = cheques.id_banco "
										l_sql = l_sql & " where cheques.empnro = " & Session("empnro")   
										
										l_sql  = l_sql  & " ORDER BY bancos.nombre_banco "
										rsOpen l_rs, cn, l_sql, 0
										do until l_rs.eof		%>	
										<option monto="<%= l_rs("importe") %>"  value= <%= l_rs("id") %> > 
										<%= l_rs("nombre_banco") %> &nbsp;-&nbsp;<%= l_rs("numero") %> </option>
										<%	l_rs.Movenext
										loop
										l_rs.Close %>
									</select>
									<script>document.datos_02.idcheque.value= "<%= l_idcheque %>"</script></div>
       
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
									<script>document.datos_02.idresponsable.value= "<%= l_idresponsable %>"</script>	</div>
       
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

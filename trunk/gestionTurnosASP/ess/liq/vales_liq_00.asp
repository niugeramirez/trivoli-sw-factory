<%Option Explicit %>
<!--#include virtual="/serviciolocal/shared/inc/sec.inc"-->
<!--#include virtual="/serviciolocal/shared/inc/const.inc"-->
<!--#include virtual="/serviciolocal/shared/db/conn_db.inc"-->
<!--#include virtual="/ess/ess/shared/inc/const.inc"-->

<!--
-----------------------------------------------------------------------------
Archivo        : vales_liq_00.asp
Descripcion    : Modulo que se encarga del ABM de vales
Creador        : Scarpa D.
Fecha Creacion : 09/01/2004
Modificacion   :
  27/01/2003 - Scarpa D. - Sacar del titulo la palabra configuracion
                           No permitir borrar/modificar vales liquidados
  10-02-2004 - F. Favre - Se agrando la llamada a la ventana vales_liq_02
  23-04-2004 - J.M. Hoffman - Se arreglo el botón orden y filtro
  07-07-2004 - Alvaro Bayon - Botón pago
  Modificado  : 12/09/2006 Raul Chinestra - se agregó Vales en Autogestión   
  30-07-2007 - Diego Rosso - Se cambio el formato de la tabla.
-----------------------------------------------------------------------------
-->
<%
dim l_rs
dim l_sql

Dim l_tvalenro
Dim l_pliqnro
Dim l_empleg

l_tvalenro  = 0
l_pliqnro   = 0

' Filtro
  Dim l_Etiquetas  ' Son los nombres que deben aparecer en la ventana para que el usuario seleccione
  Dim l_Campos     ' Son los campos de la base que apareceran en la clausula where, que deben estar asociados a las etiquetas
  Dim l_Tipos      ' Son los tipos de datos que tienen los campos (N=Numerico, T=Texto y F=Fecha)

' Orden
  Dim l_Orden      ' Son las etiquetas que aparecen en el orden
  Dim l_CamposOr   ' Son los campos para el orden
  
' Filtro
  l_etiquetas = "Código:;Descripción:;Monto:;Fec.Pedido:"
  l_Campos    = "vales.valnro;valdesc;valmonto;valfecped"
  l_Tipos     = "N;T;N;F"

' Orden
  l_Orden     = "Código:;Descripción:;Monto:;Fec.Pedido:"
  l_CamposOr  = "vales.valnro;valdesc;valmonto;valfecped"
  
  l_empleg = request("empleg")
  
%>
<html>
<head>
<link href="../<%=c_estilo %>" rel="StyleSheet" type="text/css">
<meta http-equiv="Content-Type" content="text/html; charset=iso-8859-1">
<title>Vales</title>
<script src="/serviciolocal/shared/js/fn_windows.js"></script>
<script src="/serviciolocal/shared/js/fn_confirm.js"></script>
<script src="/serviciolocal/shared/js/fn_ayuda.js"></script>
<script>
function Alta(){
  abrirVentana('vales_liq_02.asp?Tipo=A' + '&pliqnro=' + document.datos.pliqnro.value + '&tvalenro=' + document.datos.tvalenro.value +"&empleg=<%=l_empleg %>",'',550,350);
}

function Modificar(){
 if (ifrm.jsSelRow != null){
   if ( (ifrm.jsSelRow.cells(5).innerText == 'SI ') || (ifrm.jsSelRow.cells(6).innerText == 'SI') ) {
            alert('No se puede Modificar el Vale seleccionado.');
			return;
	}			
	else {		
        var param = '?Tipo=M&valnro=' + document.ifrm.datos.cabnro.value+ '&pliqnro=' + document.datos.pliqnro.value + '&tvalenro=' + document.datos.tvalenro.value + "&empleg=<%= l_empleg %>";  
        abrirVentana('vales_liq_02.asp' + param,'',550,350);
	}		
 }			
 else
    alert('Debe Seleccionar un Registro.');

}


function orden(pag)
{
  abrirVentana('/ess/ess/shared/asp/orden_param_adp_00.asp?pagina='+pag+'&lista=<%= l_orden %>&campos=<%= l_camposOr%>&filtro='+escape(document.ifrm.datos.filtro.value),'',350,160)  
  abrirVentana('/ess/ess/shared/asp/orden_param_adp_00.asp?pagina='+pag+'&lista=<%= l_orden %>&campos=<%= l_camposOr%>&filtro='+escape(document.ifrm.datos.filtro.value),'',350,160)  
}

function filtro(pag)
{
  abrirVentana('/ess/ess/shared/asp/filtro_param_adp_00.asp?pagina='+pag+'&campos=<%= l_campos%>&tipos=<%=l_tipos%>&etiquetas=<%=l_etiquetas%>&orden='+document.ifrm.datos.orden.value,'',250,160);
} 

function param(){
  return ('pliqnro=' + document.datos.pliqnro.value + '&tvalenro=' + document.datos.tvalenro.value + '&empleg=<%=l_empleg %>');
 
}

function salidaExcel(){
  abrirVentana('vales_liq_excel.asp?filtro='+ escape(document.ifrm.datos.filtro.value)  + '&'+ param()+'&orden='+ document.ifrm.datos.orden.value,'',300,300);
}    	   

function cambioTipoVale(codigo){
  document.datos.tvalenro.value = codigo;
  actualizar();
}

function  actualizar(){
  if (document.datos.pliqnro.value != ""){
     ifrm.location = "vales_liq_01.asp?empleg=<%=l_empleg%>" + '&pliqnro=' + document.datos.pliqnro.value + '&tvalenro=' + document.datos.tvalenro.value
  }else{
     ifrm.location = "blanc.html" 
  }
}

function config(){
  abrirVentana('vales_config_liq_00.asp','',330,140);
}


function baja(){

 if (ifrm.jsSelRow != null){
   if ( (ifrm.jsSelRow.cells(5).innerText.toString() == 'SI ') || (ifrm.jsSelRow.cells(6).innerText.toUpperCase() == 'SI') ) {
            alert('No se puede Eliminar el Vale seleccionado.');
			return;
	}			
	else		
      eliminarRegistro(document.ifrm,'vales_liq_04.asp?valnro='+document.ifrm.datos.cabnro.value);
 }			
 else
    alert('Debe Seleccionar un Registro.');

}

</script>

</head>
<body leftmargin="0" topmargin="0" rightmargin="0" bottommargin="0">
<form name="datos" target="vent_oculta" method="post">
<input type="Hidden" name="empleados" value=""> 
<table border="0" cellpadding="0" cellspacing="0" height="5%">
<tr >
<td align="left" valign="top">Vales</td>
<td align="right" >
    <% call MostrarBoton ("sidebtnABM", "Javascript:Alta();","Alta") %>
    <% call MostrarBoton ("sidebtnABM", "Javascript:baja();","Baja") %>	
    <% call MostrarBoton ("sidebtnABM", "Javascript:Modificar();","Modifica") %>
    &nbsp;&nbsp;&nbsp;
    <% call MostrarBoton ("sidebtnSHW", "Javascript:salidaExcel();","Excel") %>	
    &nbsp;&nbsp;&nbsp;

	<a class=sidebtnSHW href="Javascript:orden('/ess/ess/liq/vales_liq_01.asp');">Orden</a>	
 	<a class=sidebtnSHW href="Javascript:filtro('/ess/ess/liq/vales_liq_01.asp');">Filtro</a>	
    &nbsp;&nbsp;&nbsp;	
</td>	  
</tr>
<tr>
   <td align="right">
     <b>Período&nbsp;Dto.:</b>
   </td>  
   <td>			
	   <select name="pliqnro" size="1" style="width:200px" onchange="javascript:actualizar();">
		<%	Set l_rs = Server.CreateObject("ADODB.RecordSet") 
		
  		    l_sql = "SELECT pliqnro, pliqdesc "
			l_sql  = l_sql  & "FROM periodo "
			l_sql  = l_sql  & "ORDER BY pliqanio DESC, pliqmes DESC "
			rsOpen l_rs, cn, l_sql, 0
			do until l_rs.eof		%>	
			<option value="<%= l_rs("pliqnro") %>" <%if CInt(l_rs("pliqnro")) = CInt(l_pliqnro) then response.write "selected" end if%> > 
			<%= l_rs("pliqdesc") %> </option>
		<%	    l_rs.Movenext
			loop
			l_rs.Close        
			  %>	
		</select>
   </td>
</tr>
<tr>
   <td align="right">
     <b>Tipo Vale:</b>
   </td>
   <td align="left" colspan="3"> 	
	 <select name="tvalenro" size="1" onchange="javascript:actualizar();">
			<option value="">Todos</option>
		<%	Set l_rs = Server.CreateObject("ADODB.RecordSet") 
		
  		    l_sql = "SELECT tvalenro, tvaledesabr "
			l_sql  = l_sql  & "FROM tipovale "
			l_sql  = l_sql  & "ORDER BY tvaledesabr "						
			rsOpen l_rs, cn, l_sql, 0
			do until l_rs.eof		%>	
			<option value="<%= l_rs("tvalenro") %>" <%if CInt(l_rs("tvalenro")) = CInt(l_tvalenro) then response.write "selected" end if%> > 
			<%= l_rs("tvaledesabr") %> </option>
		<%	    l_rs.Movenext
			loop
			l_rs.Close %>	
	  </select>
   </td>
</tr>
</table>
<table border="0" cellpadding="0" cellspacing="0" height="95%">
<tr valign="top">
   <td colspan="2" style="" height="100%">
     <iframe name="ifrm" src="vales_liq_01.asp?empleg=<%=l_empleg%>" width="100%" height="100%"></iframe> 
   </td>
</tr>
</table>
</form>

<script>
  actualizar();
</script>
</body>
</html> 

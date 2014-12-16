<% Option Explicit %>
<!--#include virtual="/serviciolocal/shared/db/conn_db.inc"-->
<!--#include virtual="/serviciolocal/shared/inc/fecha.inc"-->
<!--
'-----------------------------------------------------------------------------------
Archivo	    : barra_proceso_00.asp
Descripción : Barra de estado del proceso
Autor		: Scarpa D. 
Fecha		: 05/02/2004
Modificado	: 
	18-02-04 Favre F. Se standariso
------------------------------------------------------------------------------------
-->
<% 
'Variables base de datos
 
'Variables uso local
 Dim l_porc
 
 Dim l_tiempo_actual
 Dim l_tiempo_restante
 Dim l_tiempo_total
 
 Dim l_bpronro
 Dim l_funcion
 Dim l_parametros
 
 l_bpronro 	  = request("bpronro")
 l_funcion	  = Request.QueryString("funcion")
 l_parametros = Request.QueryString("parametros")
 
%>
<script src="/serviciolocal/shared/js/fn_windows.js"></script>

<style>
.hidden
{
	background : transparent;
	border : none;
    COLOR: black;
	FONT-FAMILY: Verdana;
	FONT-SIZE: 08pt;
}
</style>
<html>
<head>
<title><%= Session("Titulo")%>Estado Proceso</title>
</head>
<link href="/serviciolocal/shared/css/tables_gray.css" rel="StyleSheet" type="text/css">
<body leftmargin="0" rightmargin="0" topmargin="0" bottommargin="0">

<script>
function cambiar(porcentaje,tiempoActual, tiempoRest, barra){
  document.all.tiemproc.value = tiempoActual;  
  document.all.tiemrest.value = tiempoRest;  	
  if (porcentaje == '100'){
     document.all.etiqueta.value = 'Proceso Finalizado';
  }else{
     document.all.etiqueta.value = 'Procesando...';
  }
  var td = document.getElementById("barraprog");
  td.innerHTML = barra;
}
</script>

<table width="100%" height="100%" border="0" cellpadding="0" cellspacing="0">
  <tr>
    <td colspan="3">
	  <br>
	</td>
  </tr>
  <tr>
    <td>
	   <br>
	</td> 
    <td align="center" valign="middle" style="width:320px;height:250px;">
	
		<table  width="100%" height="100%" border="0" cellpadding="0" cellspacing="0" style="border : thick none #E4FEF9;">
		  <tr>
		    <td valign="bottom" align="center" height="40%">
			  <input type="Text" name="etiqueta" size="25" class="hidden" readonly>
			</td>
		  </tr>
		  <tr>
		    <td id="barraprog" align="center" height="20%">
		      
		    </td>
		  </tr>
		  <tr>
		    <td valign="bottom" align="center" height="40%">
			   <table align="center" cellpadding="0" cellspacing="0" border="0" style="width:260px;border : thick none #E4FEF9;">
				  <tr>
				     <td colspan="2">
					   <hr width="100%" size="1">
					 </td>
				  </tr>		  
			      <tr>
				    <td align="left">
					   <b>Tiempo&nbsp;Procesado:&nbsp;</b>
					</td>		  
					<td align="right">
 			            <input type="Text" name="tiemproc" size="9" class="hidden" readonly>
					</td>
				  </tr>
			      <tr>
				    <td align="left">
		         		<b>Tiempo&nbsp;Estimado&nbsp;Restante:&nbsp;</b>	
					</td>		  
					<td align="right">
 			            <input type="Text" name="tiemrest" size="9" class="hidden" readonly>
					</td>
				  </tr>
				  <tr>
				     <td colspan="2">
					   <hr width="100%" size="1">
					 </td>
				  </tr>
			   </table>
			</td>
		  </tr>  
		</table>
	
	</td>
    <td>
	   <br>
	</td> 
  </tr>
  <tr>
    <td colspan="3">
      <br>	
	</td>
  </tr>
</table>

<iframe frameborder="0" height="0%" width="0%" scrolling="No" align="center" src="barra_proceso_01.asp?bpronro=<%= l_bpronro%>&funcion=<%= l_funcion %>&parametros=<%= l_parametros %>" ></iframe> 

</body>
</html>






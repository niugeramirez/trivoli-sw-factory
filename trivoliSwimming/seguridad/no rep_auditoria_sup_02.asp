<% Option Explicit %>
<!--#include virtual="/trivoliSwimming/shared/db/conn_db.inc"-->
<!--#include virtual="/trivoliSwimming/shared/inc/fecha.inc"-->
<!--
-----------------------------------------------------------------------------
Archivo        : rep_auditoria_sup_02
Descripcion    : Se encarga de testear si el proceso termino.
Creador        : JMH
Fecha Creacion : 20/01/2005
Modificacion   :
-----------------------------------------------------------------------------
-->
<% 
on error goto 0

'Variables base de datos
 Dim l_rs
 Dim l_cm
 Dim l_sql

'Variables uso local
 Dim l_porc
 Dim l_bpronro
 
 Dim l_tiempo_actual
 Dim l_tiempo_restante
 Dim l_tiempo_total
 Dim l_total_emp
 Dim l_restantes_emp
 
 Dim l_desde
 Dim l_hasta
 Dim l_empresa
 Dim l_incOperBen

 l_bpronro = request("bpronro")
 l_total_emp = request("totalemp")

%>
<script src="/rhprox2/shared/js/fn_windows.js"></script>

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
<title>Estado Proceso</title>
</head>
<link href="/rhprox2/shared/css/tables_gray.css" rel="StyleSheet" type="text/css">
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

<iframe frameborder="1" height="0" width="0" scrolling="No" align="center" src="rep_auditoria_sup_03.asp?bpronro=<%= l_bpronro%>" ></iframe> 

</body>
</html>






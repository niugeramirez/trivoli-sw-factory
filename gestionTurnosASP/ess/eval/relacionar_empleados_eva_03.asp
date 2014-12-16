<% Option Explicit %>
<!--#include virtual="/ess/ess/shared/inc/const.inc"-->


<!--#include virtual="/serviciolocal/shared/db/conn_db.inc"-->
<!--#include virtual="/serviciolocal/shared/inc/fecha.inc"-->
<!--#include virtual="/serviciolocal/shared/inc/const_eva.inc"-->
<!--
Archivo	    : relacionar_empleados_eva_03.asp
Descripción : Barra de estado del proceso de creacion formularios evaluacion
Autor		: CCRossi
Fecha		: 18-01-2005
Modificado	: 
-->
<% 


'Variables base de datos
 Dim l_rs
 Dim l_cm
 Dim l_sql

'Variables uso local
 Dim l_porc
 Dim l_bpronro
 Dim l_total_emp
 
 Dim l_tiempo_actual
 Dim l_tiempo_restante
 Dim l_tiempo_total
 Dim l_restantes_emp
 
 Dim l_hasta
 Dim l_empresa

 l_bpronro = request("bpronro")
 l_total_emp = request("totalemp")
 
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
<title>Estado Proceso</title>
</head>
<link href="../<%=c_estilo %>" rel="StyleSheet" type="text/css">
<body leftmargin="0" rightmargin="0" topmargin="0" bottommargin="0">

<script>
function cambiar(porcentaje,totalEmp, restEmp,tiempoActual, tiempoRest, barra){
  document.all.totemp.value = totalEmp;
  document.all.resemp.value = restEmp;  
  document.all.tiemproc.value = tiempoActual;  
  document.all.tiemrest.value = tiempoRest;  	
  if (porcentaje == '100'){
     document.all.etiqueta.value = 'Proceso Finalizado';
  }else{
<%if ccodelco=-1 then %>
     	document.all.etiqueta.value = 'Procesando Supervisados...';
<%else%>
	document.all.etiqueta.value = 'Procesando Empleados...';
<%end if%>
  }
  var td = document.getElementById("barraprog");
  td.innerHTML = barra;
}

function actualizarEmpleados(){
  //02: parent 00: opener
  opener.parent.location.reload();
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
		         		<b>Total <%if ccodelco=-1 then%>Supervisados<%else%>Empleados<%end if%>:&nbsp;</b>	
					</td>		  
					<td align="right">
 			            <input type="Text" name="totemp" size="9" class="hidden" readonly>
					</td>
				  </tr>
			      <tr>
				    <td align="left">
		         		<b><%if ccodelco=-1 then%>Supervisados<%else%>Empleados<%end if%>&nbsp;Restantes:&nbsp;</b>	
					</td>		  
					<td align="right">
 			            <input type="Text" name="resemp" size="9" class="hidden" readonly>
					</td>
				  </tr>	   
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

<iframe frameborder="1" height="0" width="0" scrolling="No" align="center" src="relacionar_empleados_eva_04.asp?bpronro=<%= l_bpronro%>&totalemp=<%= l_total_emp%>"></iframe> 

</body>
</html>



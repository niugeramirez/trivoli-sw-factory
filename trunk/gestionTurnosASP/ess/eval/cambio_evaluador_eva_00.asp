<%Option Explicit %>
<!--#include virtual="/serviciolocal/shared/inc/sec.inc"-->
<!--#include virtual="/serviciolocal/shared/inc/const.inc"-->
<!--#include virtual="/serviciolocal/shared/inc/const_eva.inc"-->
<!--#include virtual="/serviciolocal/shared/db/conn_db.inc"-->
<%
'---------------------------------------------------------------------------------------
'Archivo	: equipo_eva_00.asp
'Descripción: ABM de equipo de trabajo del proyecto
'Autor		: CCRossi
'Fecha		: 16-12-2004
'Modificado : 02-04-2005 - LAmadio - funcione con VB - y generac form eva.
'---------------------------------------------------------------------------------------
 on error goto 0
 
' Variables
' de parametros entrada
  dim l_evaproynro
  dim l_ternro ' ternro del gerente que está logeado ( o socio)
  
   dim l_evaproynom 
   dim l_evaevenro
  '  evaevenro-...---------
  
  
' de uso local  
  Dim l_seleccion 
  Dim l_listempleados    
  dim l_listinicial
  dim l_perfil
		' evaevedesabr ---------
  
' de base de datos  
  Dim l_sql
  Dim l_rs

' Filtro
  Dim l_Etiquetas  ' Son los nombres que deben aparecer en la ventana para que el usuario seleccione
  Dim l_Campos     ' Son los campos de la base que apareceran en la clausula where, que deben estar asociados a las etiquetas
  Dim l_Tipos      ' Son los tipos de datos que tienen los campos (N=Numerico, T=Texto y F=Fecha)

' Orden
  Dim l_Orden      ' Son las etiquetas que aparecen en el orden
  Dim l_CamposOr   ' Son los campos para el orden

' Filtro
  l_etiquetas = "Empleado:;Apellido y Nombre:"
  l_Campos    = "empleado.empleg;empleado.terape"
  l_Tipos     = "N;T"

' Orden
  l_Orden     = "Empleado:;Apellido y Nombre:"
  l_CamposOr  = "empleado.empleg;empleado.terape"


  
		' request evento --------
 l_evaproynro = request("evaproynro")
 
 l_evaevenro=""   
 if (l_evaproynro<>"" and not isnull(l_evaproynro)) then
	Set l_rs  = Server.CreateObject("ADODB.RecordSet")
	l_sql = "SELECT evaproyecto.evaproynro, evaproynom, evaevento.evaevenro  "
	l_sql = l_sql & " FROM  evaproyecto "
	l_sql = l_sql & " INNER JOIN evaevento ON evaevento.evaproynro=evaproyecto.evaproynro "
	l_sql = l_sql & " WHERE evaproyecto.evaproynro=" & l_evaproynro
	rsOpen l_rs, cn, l_sql, 0 
	if not l_rs.EOF then
		l_evaproynom= l_rs("evaproynom")
		l_evaevenro= l_rs("evaevenro")
	end if	
	l_rs.close
	set l_rs =nothing
 end if
  
'cargar empleados ya asociados al PROYECTO.... para cuando entre en el select.
	' para RDE verrrrrrrr para RDP????????
 l_listempleados = "0"
 Set l_rs = Server.CreateObject("ADODB.RecordSet")
 l_sql = "SELECT DISTINCT evaproyemp.ternro "
 l_sql = l_sql & " FROM  evaproyemp "
 l_sql = l_sql & " WHERE  evaproyemp.evaproynro= " & l_evaproynro
 rsOpen l_rs, cn, l_sql, 0 
 do until l_rs.eof 
 	l_listempleados = l_listempleados & "," & l_rs("ternro")
	
'' OJO(en relac empls) - l_listempleados = l_listempleados & "," & l_rs("empleado") &"@"& l_rs("empleg")
	l_rs.MoveNext
 loop
 l_rs.close
 set l_rs = nothing
 l_listinicial = l_listempleados ' lista inicial de ternros 
%>

<html>
<head>
<link href="/serviciolocal/shared/css/tables3.css" rel="StyleSheet" type="text/css">
<meta http-equiv="Content-Type" content="text/html; charset=iso-8859-1">
<title>Empleados del Proyecto - Gesti&oacute;n de Desempeño - RHPro &reg;</title>
<script src="/serviciolocal/shared/js/fn_windows.js"></script>
<script src="/serviciolocal/shared/js/fn_confirm.js"></script>
<script src="/serviciolocal/shared/js/fn_ayuda.js"></script>
<script>
var grabo;
grabo=0;

function filtro(pag)
{
  abrirVentana('filtro_param_adp_00.asp?pagina='+pag+'&campos=<%= l_campos%>&tipos=<%=l_tipos%>&etiquetas=<%=l_etiquetas%>&orden='+document.ifrm.datos.orden.value,'',250,160);
}

function orden(pag)
{
abrirVentana('orden_param_adp_00.asp?pagina='+pag+'&lista=<%= l_orden %>&campos=<%= l_camposOr%>&filtro='+escape(document.ifrm.datos.filtro.value),'',350,160)
}
 	   
function param(){
	var chequear;
	chequear=  'evaproynro=<%=l_evaproynro%>';
	return chequear;
}

function abrirVentanaVerif(url, name, width, height, opc, sistema){
  if (ifrm.jsSelRow == null)
    alert("Debe seleccionar un registro.")
  else	
    if ((sistema != null) &&
        (ifrm.jsSelRow.cells(sistema).innerText.toUpperCase() == 'SISTEMA'))
       alert('Registro del sistema. No se lo puede modificar.');
	else   
       abrirVentana(url, url.substr(0,url.indexOf(".asp")), width, height, opc) 
}

function eliminarRegistro(obj,donde,sistema)
{
	if (obj.datos.cabnro.value == 0)
		{
		alert("Debe selecionar un registro para realizar la operacion.");
		}
	else
		{
        if ((sistema != null) &&
            (ifrm.jsSelRow.cells(sistema).innerText.toUpperCase() == 'SISTEMA'))
            alert('Registro del sistema. No se lo puede Borrar.');
		else
			if (confirm('¿ Desea eliminar el registro seleccionado ?') == true)
 				{
				abrirVentanaH(donde,"",250,120);
  				}
		}
}

function llamadaexcel(){ 
	/* if (filtro == "")
		Filtro(true);
	else
		abrirVentana("equipo_eva_excel.asp?orden=" + document.ifrm.datos.orden.value + "&filtro=" + escape(document.ifrm.datos.filtro.value)+'&evaproynro=<%=l_evaproynro%>&listempleados='+ document.datos.listempleados.value,'execl',250,150);
*/
}


function llamargengrup(){
	 grabo=0;
	 abrirVentana('../shared/asp/gen_select_emp_00.asp?llamadora=EVA&seltipnro=-1&srcdatos=opener.document.all.listempleados&filtroini=select empleado.ternro from empleado   inner join evaproyemp on   evaproyemp.ternro=empleado.ternro and evaproyemp.evaproynro=<%=l_evaproynro%>','',700,570);	  
	 
	//  o  ../shared/asp/gen_select_emp_v2_00.asp
}

function actualizarEmpleados(){
//alert(document.datos.listempleados.value);
// document.ifrm.location = '/serviciolocal/eval/equipo_eva_01.asp?evaproynro=<%= l_evaproynro%>&listempleados=' + document.datos.listempleados.value;
			// en relac empls no pasa lista empls.....!!1
// document.ifrm.location = '../../eval/equipo_eva_01.asp?evaproynro=<%= l_evaproynro%>&listempleados=' + document.datos.listempleados.value;
}

var VentProcesando;

function Validar_Formulario(){
	//abrirVentana('equipo_eva_02.asp?evaproynro=<%'=l_evaproynro%>&listempleados='+ document.datos.listempleados.value+'&listinicial='+ document.datos.listinicial.value,'',300,150);	  
	//grabo=1;
	
	
  //abrirVentanaH('','voculta',300,300);
  //document.datos.target = 'voculta';
  //document.datos.action ='equipo_eva_02.asp?evaproynro=<%=l_evaproynro%>' //&listempleados='+ document.datos.listempleados.value+'&listinicial='+ document.datos.listinicial.value
//  document.datos.action = '/serviciolocal/eval/relacionar_empleados_eva_02.asp?evaevenro=<%'= l_evaevenro%>';
 // document.datos.submit();	 
  
  //abrirVentanaH('equipo_eva_02.asp?evaproynro=<%=l_evaproynro%>&listempleados='+ document.datos.listempleados.value+'&listinicial='+ document.datos.listinicial.value,"",250,120);
    
  //grabo=1;
}

function Chequear(){
	if ((document.datos.listempleados.value!==document.datos.listinicial.value) && (grabo==0))
	{
		if (confirm('¿ Confirma Aceptar sin Grabar ?') == true)
 			window.close();
 	}		
 	else		
 		window.close();
}

function CerrarVentana(){
	VentProcesando.close();  
}  
</script>

</head>
<body leftmargin="0" topmargin="0" rightmargin="0" bottommargin="0">
<form name="datos"> <!-- form datos 1??????? -->
<input type="hidden" name="seleccion">
<input type="hidden" name="listempleados" value="<%=l_listempleados%>">
<input type="hidden" name="listinicial"	  value="<%=l_listinicial%>">

<table border="0" cellpadding="0" cellspacing="0" height="95%">
<tr style="border-color :CadetBlue;">
	<td align="left" class="barra">Empleados del Proyecto:&nbsp;<%=l_evaproynom%></td> <!-- Empleados del Proyecto: -->
	<td nowrap align="right" class="barra">
		<a class=sidebtnABM href="Javascript:abrirVentanaVerif('evaluadores_DEL_eva_00.asp?ternro=' + document.ifrm.datos.cabnro.value+'&evaevenro=<%=l_evaevenro%>','',500,270)">Evaluadores</a>
		&nbsp;&nbsp;
		<!-- 		<a class=sidebtnSHW href="Javascript:llamadaexcel();">Excel</a>  -->
		<a class=sidebtnSHW href="Javascript:orden('cambio_evaluador_eva_01.asp');">Orden</a>
		<a class=sidebtnSHW href="Javascript:filtro('cambio_evaluador_eva_01.asp');">Filtro</a>
		<a class=sidebtnHLP href="Javascript:ayuda('<%= Request.ServerVariables("SCRIPT_NAME")%>');">Ayuda</a>
		<br><br>
		<input type="hidden" NAME="veranterior" value="0">
		<!-- 	<a class=sidebtnABM href="Javascript:llamargengrup();">Incorporar&nbsp; Empleados&nbsp;</a>		-->
		<%			'call MostrarBoton ("sidebtnABM", "Javascript:abrirVentanaVerif('evaluadores_eva_00.asp?Tipo=M&evaproynro=' + document.ifrm.datos.cabnro.value+'&perfil="&l_perfil&"','',300,300);","Cambiar Revisor")		%>
		<!--INPUT onclick="Javascript:llamargengrup()" type=button value="Relacionar" name=Gengrup--> 
	</td>
</tr>
<tr valign="top" height="100%">
   <td colspan="2">
   <iframe name="ifrm" src="cambio_evaluador_eva_01.asp?evaproynro=<%=l_evaproynro%>" width="100%" height="100%"></iframe> 
   </td>
</tr>
<tr>
	<td colspan="2" height="10">
	</td>
</tr>
<tr>
    <td colspan=2 align="right" class="th2">
	<!-- 		<a class=sidebtnABM href="Javascript:Validar_Formulario()">Grabar</a>  -->
		<!-- Javascript:Chequear()-->
		<a class=sidebtnABM href="Javascript:window.close()">Aceptar</a>
		<a class=sidebtnABM href="Javascript:window.close()">Cancelar</a>
	</td>
</tr>
</table>
</form>	
</body>
</html>

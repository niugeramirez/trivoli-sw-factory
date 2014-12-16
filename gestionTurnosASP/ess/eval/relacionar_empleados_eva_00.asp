<%Option Explicit %>
<!--#include virtual="/serviciolocal/ess/shared/inc/const.inc"-->

<!--#include virtual="/serviciolocal/shared/inc/sec.inc"-->
<!--#include virtual="/serviciolocal/shared/inc/const.inc"-->
<!--#include virtual="/serviciolocal/shared/inc/const_eva.inc"-->
<!--#include virtual="/serviciolocal/shared/db/conn_db.inc"-->
<%
'---------------------------------------------------------------------------------------
'Archivo	: relacionar_empleados_eva_00.asp
'Descripción	: ABM de relacion empleados al evento
'Autor		: CCRossi
'Fecha		: 18-05-2004
'Modificado	: 1 de Junio CCRossi. Separar en listas pequeñas de 500 empleados.

'---------------------------------------------------------------------------------------


' Variables
' de parametros entrada
  dim l_evaevenro
  
' de uso local  
  Dim l_seleccion 
  Dim l_listempleados   'Todos para el gen_select
  dim l_listinicial     'Todos los que ya estaban al cargar la pagina

  dim l_evaevedesabr
  
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
if ccodelco=-1 then
  l_etiquetas = "N&uacute;mero:;Apellido:"
  l_Orden     = "N&uacute;mero:;Apellido:"
else
  l_etiquetas = "Empleado:;Apellido:"
  l_Orden     = "Empleado:;Apellido:"
end if
  l_Campos    = "empleado.empleg;empleado.terape"
  l_Tipos     = "N;T"

' Orden
  l_CamposOr  = "empleado.empleg;empleado.terape"

 l_evaevenro = request("evaevenro")
    
 if (l_evaevenro<>"" and not isnull(l_evaevenro)) then
	Set l_rs  = Server.CreateObject("ADODB.RecordSet")
	l_sql = "SELECT evaevento.evaevenro, evaevento.evaevedesabr FROM  evaevento WHERE evaevenro=" & l_evaevenro
	rsOpen l_rs, cn, l_sql, 0 
	if not l_rs.EOF then
		l_evaevedesabr= l_rs("evaevedesabr")
	end if	
	l_rs.close
	set l_rs =nothing
 end if
  
'cargar empleados ya asociados al evento.... para cuando entre en el select.
 l_listempleados  = "0"
 
 Set l_rs = Server.CreateObject("ADODB.RecordSet")
 l_sql = "SELECT DISTINCT evacab.empleado, empleado.empleg FROM  evacab INNER JOIN empleado ON empleado.ternro = evacab.empleado WHERE  evacab.evaevenro   = " & l_evaevenro
 l_sql = l_sql & " order by empleado.empleg "
 rsOpen l_rs, cn, l_sql, 0 
 do until l_rs.eof
'response.write l_rs("empleado") & "<br>"
	l_listempleados = l_listempleados & "," & l_rs("empleado") &"@"& l_rs("empleg")
		
	l_rs.MoveNext

 loop
 l_rs.close
 set l_rs = nothing
 
 
 l_listinicial   = l_listempleados
 
'response.write l_listempleados5 & "<br>"

%>

<html>
<head>
<link href="../<%=c_estilo %>" rel="StyleSheet" type="text/css">
<meta http-equiv="Content-Type" content="text/html; charset=iso-8859-1">
<title><%if ccodelco=-1 then%>Supervisados <%else%>Empleados <%end if%>del Evento - Gesti&oacute;n de Desempeño - RHPro &reg;</title>
<script src="/serviciolocal/shared/js/fn_windows.js"></script>
<script src="/serviciolocal/shared/js/fn_confirm.js"></script>
<script src="/serviciolocal/shared/js/fn_ayuda.js"></script>
<script>
var grabo;
grabo=0;

function filtro(pag)
{
    abrirVentana('filtro_param_eva_00.asp?llamadora=RELACIONAR&pagina='+pag+'&campos=<%= l_campos%>&tipos=<%=l_tipos%>&etiquetas=<%=l_etiquetas%>&orden='+document.ifrm.datos.orden.value,'',250,160);
}
function filtroinicial(pag)
{
	 abrirVentana('filtro_param_eva_00.asp?llamadora=RELACIONAR&pagina='+pag+'&campos=<%= l_campos%>&tipos=<%=l_tipos%>&etiquetas=<%=l_etiquetas%>&orden='+document.ifrm.datos.orden.value,'',250,160);
}

function orden(pag)
{
abrirVentana('orden_param_eva_00.asp?pagina='+pag+'&lista=<%= l_orden %>&campos=<%= l_camposOr%>&filtro='+escape(document.ifrm.datos.filtro.value),'',350,160)
}
 	   
function param(){
	var chequear;
	chequear=  'evaevenro=<%=l_evaevenro%>&listempleados=';
	return chequear;
}
function abrirVentanaVerif(url, name, width, height, opc, sistema) 
{
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
		alert("Debe selecionar un registro para realizar la operación.");
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
	if (filtro == "")
		Filtro(true);
	else
		abrirVentana("relacionar_empleados_eva_excel.asp?orden=" + document.ifrm.datos.orden.value + "&filtro=" + escape(document.ifrm.datos.filtro.value)+'&evaevenro=<%=l_evaevenro%>&listempleados=','execl',250,150);
}

function llamargengrup(){
	 grabo=0;
	 abrirVentana('../shared/asp/gen_select_emp_v2_00.asp?llamadora=EVA&seltipnro=-1&srcdatos=opener.document.all.listempleados&filtroini=select empleado.ternro from empleado   inner join his_estructura on   his_estructura.ternro=empleado.ternro and his_estructura.htethasta is null inner join evatip_estr on   evatip_estr.estrnro=his_estructura.estrnro  inner join evaevento on   evaevento.evatipnro= evatip_estr.evatipnro and evaevento.evaevenro=<%=l_evaevenro%>','',700,570);	  
}

function actualizarEmpleados(){
 document.datos.target = 'ifrm';
 document.datos.action = 'relacionar_empleados_eva_01.asp?evaevenro=<%= l_evaevenro%>&datos=SI&filtro='+escape(document.ifrm.datos.filtro.value)+'&orden='+document.ifrm.datos.orden.value;
 document.datos.submit();	   
}

var VentProcesando;

function Validar_Formulario()
{
  abrirVentanaH('','voculta',300,300);
  document.datos.target = 'voculta';
  document.datos.action = 'relacionar_empleados_eva_02.asp?evaevenro=<%= l_evaevenro%>';
  document.datos.submit();	   
  grabo=1;
}

function Chequear()
{
	if ((document.datos.listempleados.value!==document.datos1.listinicial.value) && (grabo==0))
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
<body leftmargin="0" topmargin="0" rightmargin="0" bottommargin="0" onload="filtroinicial('relacionar_empleados_eva_01.asp');">

<form name="datos1" method="post">
<input type="hidden"   name="listinicial"    value="<%=l_listinicial%>">
</form>

<form name="datos" method="post">
<input type="hidden" name="seleccion">
<input type="hidden" name="listempleados"  value="<%=l_listempleados%>">
<input type="hidden" name="datos" value="SI">

<table border="0" cellpadding="0" cellspacing="0" height="95%">
<tr style="border-color :CadetBlue;">
	<td align="left" class="barra"><%if ccodelco=-1 then%>Supervisados<%else%>Empleados<%end if%> del Evento:&nbsp;<%=l_evaevedesabr%></td>
	<td nowrap align="right" class="barra">
		<a class=sidebtnSHW href="Javascript:llamadaexcel();">Excel</a>
		<a class=sidebtnSHW href="Javascript:orden('relacionar_empleados_eva_01.asp');">Orden</a>
		<a class=sidebtnSHW href="Javascript:filtro('relacionar_empleados_eva_01.asp');">Filtro</a>
		<a class=sidebtnHLP href="Javascript:ayuda('<%= Request.ServerVariables("SCRIPT_NAME")%>');">Ayuda</a>
		<br><br>
		<input type="hidden" NAME="veranterior" value="0">
		<a class=sidebtnABM href="Javascript:llamargengrup();">Relacionar</a>
		<a class=sidebtnABM href="Javascript:abrirVentanaVerif('evaluadores_eva_00.asp?ternro=' + document.ifrm.datos.cabnro.value+'&evaevenro=<%=l_evaevenro%>','',460,425)"><%if ccodelco=-1 then%>Roles<%else%>Evaluadores<%end if%></a>
	</td>
</tr>
</form>	

<tr valign="top" height="100%">
   <td colspan="2">
   <iframe name="ifrm" src="relacionar_empleados_eva_01.asp?evaevenro=<%=l_evaevenro%>&datos=NO&Inicio=SI" width="100%" height="100%"></iframe> 
   </td>
</tr>
<tr>
	<td colspan="2" height="10">
	</td>
</tr>
<tr>
    <td colspan=2 align="right" class="th2">
		<a class=sidebtnABM href="Javascript:Validar_Formulario()">Grabar</a>
		<a class=sidebtnABM href="Javascript:Chequear()">Aceptar</a>
		<a class=sidebtnABM href="Javascript:window.close()">Cancelar</a>
	</td>
</tr>
</table>

</body>
</html>

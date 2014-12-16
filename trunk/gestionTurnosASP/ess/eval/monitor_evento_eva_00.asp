<%Option Explicit %>
<!--#include virtual="/ess/ess/shared/inc/const.inc"-->
<!--#include virtual="/serviciolocal/shared/inc/sec.inc"-->
<!--#include virtual="/serviciolocal/shared/inc/const.inc"-->
<!--#include virtual="/serviciolocal/shared/inc/const_eva.inc"-->
<!--#include virtual="/serviciolocal/shared/db/conn_db.inc"-->
<%
'Archivo	: monitor_evento_eva_00.asp
'Descripción: Ver formularios aprobados y no aprobados.
'Autor		: CCRossi
'Fecha		: 22-07-2004
'Modificado	: 
'--------------------------------------------------------------------------------------
'Variables
'Filtro
 Dim l_Etiquetas  ' Son los nombres que deben aparecer en la ventana para que el usuario seleccione
 Dim l_Campos     ' Son los campos de la base que apareceran en la clausula where, que deben estar asociados a las etiquetas
 Dim l_Tipos      ' Son los tipos de datos que tienen los campos (N=Numerico, T=Texto y F=Fecha)

'Orden
 Dim l_Orden      ' Son las etiquetas que aparecen en el orden
 Dim l_CamposOr   ' Son los campos para el orden

'Filtro
if ccodelco=-1 then
 l_etiquetas = "N&uacute;mero:;Apellido y Nombre:;Rol:"
 l_Orden     = "N&uacute;mero:;Apellido y Nombre:;Rol:;"
else
 l_etiquetas = "Empleado:;Apellido y Nombre:;Rol:"
 l_Orden     = "Empleado:;Apellido y Nombre:;Rol:;"
end if 
 l_Campos    = "empleado.empleg;empleado.terape;evaluador.terape"
 l_Tipos     = "T;T;T"

'Orden

 l_CamposOr  = "empleado.empleg;empleado.terape;evaluador.terape"

 Dim l_sql
 Dim l_rs

 Dim l_evaevenro
 l_evaevenro= Request.QueryString("evaevenro")
 
 %>

<html>
<head>
<link href="../<%=c_estilo %>" rel="StyleSheet" type="text/css">
<meta http-equiv="Content-Type" content="text/html; charset=iso-8859-1">
<title>Monitor de Eventos&nbsp;-&nbsp;Gesti&oacute;n de Desempeño - RHPro &reg;</title>
<script src="/serviciolocal/shared/js/fn_windows.js"></script>
<script src="/serviciolocal/shared/js/fn_confirm.js"></script>
<script src="/serviciolocal/shared/js/fn_ayuda.js"></script>

<script>

function filtro(pag)
{
  abrirVentana('filtro_param_adp_00.asp?pagina='+pag+'&campos=<%= l_campos%>&tipos=<%=l_tipos%>&etiquetas=<%=l_etiquetas%>&orden='+document.ifrm.datos.orden.value,'',250,160);
}

function orden(pag)
{
  abrirVentana('orden_param_adp_00.asp?pagina='+pag+'&lista=<%= l_orden %>&campos=<%= l_camposOr%>&filtro='+escape(document.ifrm.datos.filtro.value),'',350,160)
}

function Refrescar()
{
//alert('Evento =<%=l_evaevenro%>\nEstado (0:Aprobado - 1:NoAprobado - 2:Ambos)= '+ document.datos.aprobado.value+'\nMostrar (1:Resumen - 0:Detalle)= ' + document.datos.mostrar.value+'\nSecciones (-1:Obligatorias - 0:NO Obligatorias - 2:Ambos)= ' + document.datos.evaoblig.value+'\nSupervisores (0:NoHab - 1:HabNoIng - 2:IngNoTerm - 3:Todos)= ' + document.datos.control.value);
//alert('filtro=' + document.ifrm.datos.filtro.value);
//alert('listternro='+document.datos.listemp.value);
document.ifrm.location.href= 'monitor_evento_eva_01.asp?estado=' + document.datos.aprobado.value+'&evaoblig=' + document.datos.evaoblig.value+'&control=' + document.datos.control.value+'&mostrar=' + document.datos.mostrar.value+'&evaevenro=<%=l_evaevenro%>'+ "&filtro=" + escape(document.ifrm.datos.filtro.value)+'&listternro='+document.datos.listemp.value;
}

function param(){
	if (document.datos.aprobado.value=="")
		document.datos.aprobado.value=0;
	if (document.datos.mostrar.value=="")
		document.datos.mostrar.value=0;
	chequear= "estado="+document.datos.aprobado.value+"&evaoblig="+document.datos.evaoblig.value+"&control="+document.datos.control.value+"&mostrar="+document.datos.mostrar.value+"&evaevenro=<%=l_evaevenro%>";
	return chequear;
}
function llamadaexcel(){ 
	if (filtro == "")
		Filtro(true);
	else
		abrirVentana('monitor_evento_eva_excel.asp?estado=' + document.datos.aprobado.value+'&evaoblig=' + document.datos.evaoblig.value+'&control=' + document.datos.control.value+'&mostrar=' + document.datos.mostrar.value+'&evaevenro=<%=l_evaevenro%>&filtro=' + escape(document.ifrm.datos.filtro.value)+'&listternro='+document.datos.listemp.value,'excel',350,250);
}

function mail(){ 
	if (document.ifrm.datos.listamail.value=="")
		alert('No hay Evaluadores para el filtro seleccionado a quien enviarles el Mail.');
	else	
	{
		//alert('Se está cambiando la manera de enviar mails en el Sistema.\n\nEn este momento esta opción está inhabilitada.')
		abrirVentana("monitor_evento_mail_eva_00.asp?listamail=" + document.ifrm.datos.listamail.value,'',600,215);
	}	
}
 	   
 	   
function radioclick(i){
if (i==1){
	if (document.datos.quien.value!="estructura"){
		document.datos.quien.value="estructura";
		document.clasifica.location="rep_emp_por_eva_00.asp";
	}	
}
else
if (i==3){
	if (document.datos.quien.value!="empleado"){
		document.datos.quien.value="empleado";
		document.clasifica.location="rep_emp_por_eva_01.asp";
	}	
}
else{
	if (document.datos.quien.value!="evaluador"){
		document.datos.quien.value="evaluador";
		document.clasifica.location="rep_emp_por_eva_02.asp";
	}	
}	
}

function validarestructura(){
	if (document.clasifica.datos.estrnro1.value==0)
		return false;
	else
		return true;
}

function armarjoin(){
var l_join='';
if (document.clasifica.datos.estrnro1.value!=0){
	//titulo='Para las estructuras ';
	//titulo= titulo + document.clasifica.datos.nivel1.value;
	l_join= " INNER JOIN his_estructura estr1  ON  (evacab.empleado = estr1.ternro and estr1.estrnro="
			+document.clasifica.datos.estrnro1.value + ") ";
	if (document.clasifica.datos.estrnro2.value!=0){
		//titulo= titulo + ', ' + document.clasifica.datos.nivel2.value;
		l_join= l_join + " INNER JOIN his_estructura estr2 ON  (evacab.empleado = estr2.ternro and estr2.estrnro="
				+document.clasifica.datos.estrnro2.value + ") ";
		if (document.clasifica.datos.estrnro3.value!=0){
			//titulo= titulo + ', ' + document.clasifica.datos.nivel2.value;
			l_join= l_join + " INNER JOIN his_estructura estr3 ON  (evacab.empleado = estr3.ternro and estr3.estrnro="
					+document.clasifica.datos.estrnro3.value + ") ";
		}
	}
}
return l_join;
}


function cargarlistemp(){
document.datos.listemp.value="";

if (document.datos.quien.value=="estructura"){
	if (document.clasifica.datos.check.checked){
		//titulo="Todos los Evaluados";
		document.datos.listemp.value= Nuevo_Dialogo(window, 'rep_resul_totales_eva_02.asp?evaevenro='+document.datos.evaevenro.value, 500,500);
		}
	else
		if (validarestructura()){ 
			var l_join = armarjoin();
			document.datos.listemp.value= Nuevo_Dialogo(window, 'rep_resul_totales_eva_05.asp?evaevenro='+document.datos.evaevenro.value+'&join='+ l_join, 20,10);
		}	
}
else
if (document.datos.quien.value=="evaluador"){
	if ((document.clasifica.datos.ternro.value!="")&&(document.datos.evaevenro.value!="0")){
		//titulo= "Para el Evaluador "+document.clasifica.datos.empleg.value+'  '+document.clasifica.datos.empleado.value;
		document.datos.listemp.value= Nuevo_Dialogo(window, 'rep_resul_totales_eva_04.asp?evaevenro='+document.datos.evaevenro.value+'&evaluador='+ document.clasifica.datos.ternro.value, 20,10);
	}	
}
else{
	if (document.clasifica.datos.ternro.value!=""){
		//titulo="";
		document.datos.listemp.value= document.clasifica.datos.ternro.value;	
	}	
}		
}

function aplicar(){
if (document.datos.evaevenro.value=="0")
	alert("Debe seleccionar un Evento.");
else
{
	cargarlistemp();
	if (document.datos.listemp.value=="")	
	{
		<%if ccodelco=-1 then%>
		alert("No hay empleados para esta selección.");
		<%else%>
		alert("No hay supervisados para esta selección.");		
		<%end if%>
		document.datos.listemp.value="0"
	}	
	Refrescar();
	
}	
}

function generar(tipo){
	
   if (filtro != ""){
	   document.all.tipografico.value = tipo;
	   document.ifrm.generar(); 
   }
}
</script>

</head>
<body leftmargin="0" topmargin="0" rightmargin="0" bottommargin="0" onload="document.datos.aprobado.value=-1;document.datos.evaoblig.value=2;document.datos.control.value=3;document.datos.mostrar.value=1;document.datos.tipografico.value=0;">

<form name=datos>
<input type="hidden" name="tipografico" value="0">
<input TYPE="hidden" NAME="aprobado">
<input TYPE="hidden" NAME="evaoblig">
<input TYPE="hidden" NAME="control">
<input TYPE="hidden" NAME="mostrar">
<input TYPE="hidden" NAME="evaevenro" value="<%=l_evaevenro%>">

<input type="Hidden" name="ternro" value="">
<input type="hidden" name="listemp" value="">

<table border="0" cellpadding="0" cellspacing="0" height="100%">
<tr style="border-color :CadetBlue;">
	<td colspan="2" align="left" class="barra">Monitor de Eventos</td>
</tr>
<tr>
	<td colspan="2" align="right" class="barra">
	<!--a class=sidebtnSHW href="Javascript:orden('monitor_evento_eva_01.asp');">Orden</a-->
 	<a class=sidebtnSHW href="Javascript:filtro('monitor_evento_eva_01.asp');">Filtro</a>
	<a class=sidebtnSHW href="Javascript:mail();">Mail</a>
    <a class=sidebtnSHW href="Javascript:llamadaexcel();">Salida Excel</a>
	<a class=sidebtnHLP href="Javascript:ayuda('<%= Request.ServerVariables("SCRIPT_NAME")%>');">Ayuda</a>

</td>
</tr>


<tr>
<td>
<table>
<tr>
<td>
	<table>
	<tr valign="top">
   	<td align="center"  style="background-color:#FFFF00"><b>NO HABILITADO</b></td>
   	<td align="center"  style="background-color:orange"><b>INGRESO Y NO TERMINO</b></td>
   	<td align="center"  style="background-color:#FF0000"><b>HABILITADO Y NO INGRESO</b></td> 
   	<td align="center"  style="background-color:#00FF00"><b>TERMINO</b></td>  
	</tr>
	</table>
	</td>
</tr>
<tr valign="top">
   <td align="center" colspan="2"><b>Formularios:</b>
    <input TYPE="radio" NAME="estado" VALUE="0" CHECKED onclick="document.datos.aprobado.value=-1;Refrescar();">Aprobados
	<input TYPE="radio" NAME="estado" VALUE="1" onclick="document.datos.aprobado.value=0;Refrescar();">No Aprobados
	<input TYPE="radio" NAME="estado" VALUE="2" onclick="document.datos.aprobado.value=2;Refrescar();">Ambos
   </td>
</tr>
<tr valign="top">
   <td align="center" colspan="2"><b>Mostrar:</b>
    <input TYPE="radio" NAME="detalle" VALUE="0" CHECKED onclick="document.datos.mostrar.value=-1;Refrescar();">Resumen
	<input TYPE="radio" NAME="detalle" VALUE="1" onclick="document.datos.mostrar.value=0;Refrescar();">Detalle
   </td>
</tr>
<tr valign="top">
   <td align="center" colspan="2"><b>Secciones:</b>
    <input TYPE="radio" NAME="obligatoria" VALUE="0" onclick="document.datos.evaoblig.value=-1;Refrescar();">Obligatorias
	<input TYPE="radio" NAME="obligatoria" VALUE="1" onclick="document.datos.evaoblig.value=0;Refrescar();">No Obligatorias
	<input TYPE="radio" NAME="obligatoria" VALUE="2" CHECKED onclick="document.datos.evaoblig.value=2;Refrescar();">Ambas
   </td>
</tr>
<tr valign="top">
   <td align="center" colspan="2"><b><%if ccodelco=-1 then%>Supervisores:<%else%>Evaluadores:<%end if%></b>
    <input TYPE="radio" NAME="tipo" VALUE="0" onclick="document.datos.control.value=0;Refrescar();">No Habilitados
	<input TYPE="radio" NAME="tipo" VALUE="1" onclick="document.datos.control.value=1;Refrescar();">Habilitados y No Ingresaron
	<input TYPE="radio" NAME="tipo" VALUE="2" onclick="document.datos.control.value=2;Refrescar();">Ingresaron y No Terminaron
	<input TYPE="radio" NAME="tipo" VALUE="3" CHECKED onclick="document.datos.control.value=3;Refrescar();">Todos
   </td>
</tr>
</table>
</td>

<td>
	<table>
	<tr>
	<td>
		<table cellspacing="0" cellpadding="0" border="0" width="100%">
		 <tr>
			<td>
				<input type="hidden" name="quien" value="estructura">
				<input type="radio" name="quien_"  value="estructura"  onclick="Javascript:radioclick(1);" checked>
				Por<br>Estructura 
			</td>
			<td>
				<input type="radio" name="quien_"  value="evaluador"  onclick="Javascript:radioclick(2);">
				Por<br><%if ccodelco=-1 then%>Supervisor<%else%>Evaluador<%end if%>
			</td>
			<td>
				<input type="radio" name="quien_"  value="empleado"  onclick="Javascript:radioclick(3);">
				Por<br><%if ccodelco=-1 then%>Supervisado<%else%>Empleado<%end if%>
			</td>
			<td>
				<a class=sidebtnSHW href="Javascript:aplicar();">Aplicar</a>
			</td>
		</tr>
		<tr>
		<td colspan=4>
			<iframe name="clasifica" src="rep_emp_por_eva_00.asp" width="380" height="112" scrolling="No"></iframe> 
		</td>
		</tr>
		</table>
	</td>
	</tr>
	</table>
</td>

</tr>
<tr>
    <td colspan="2">
    <iframe src="../shared/asp/grafmenu.asp?graficos=0;1;3;4" height="30" width="100%" scrolling="No" frameborder="0"></iframe>
    </td>
</tr>

<tr valign="top" height="100%">
   <td colspan="2" style="">
   <iframe name="ifrm" src="monitor_evento_eva_01.asp?evaevenro=<%=l_evaevenro%>&estado=-1" width="100%" height="100%"></iframe> 
   </td>
</tr>
<tr>
	<td colspan="2" height="10"></td>
</tr>
</table>
</form>

</body>
</html>
 
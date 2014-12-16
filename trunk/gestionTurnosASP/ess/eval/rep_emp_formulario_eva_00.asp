<% Option Explicit %>
<!--#include virtual="/ess/ess/shared/inc/const.inc"-->
<!--#include virtual="/serviciolocal/shared/db/conn_db.inc"-->
<!--#include virtual="/serviciolocal/shared/inc/fecha.inc"-->
<!--#include virtual="/serviciolocal/shared/inc/const_eva.inc"-->
<% 
'---------------------------------------------------------------------------------
'Archivo	: rep_emp_formulario_eva_00.asp
'Descripción:     
'Autor		:     
'Fecha		:     
'Modificado	: 04-08-2003-Cambiar uso de estruct_actual por His_estructura
'Modificado	: 28-04-2005-CCRossi - si viene de autogestion filtrar eventos
'Modificado	: 18-05-2005 CCRossi - Traer parametros pero no lista de empleados
'            14-10-2005 - Leticia Amadio -  Adecuacion a Autogestion
'----------------------------------------------------------------------------------
on error goto 0
Dim l_rs1
Dim l_rs
Dim l_sql

Dim l_evaevenro
Dim l_tenro
Dim l_estrnro

Dim l_llamadora
Dim l_logeadoternro

l_llamadora		=Request.QueryString("llamadora")
l_logeadoternro =Request.QueryString("logeadoternro")
' si viene de AUTOGESTION trae el evaevenro.
l_evaevenro     =Request.QueryString("evaevenro")
%>

<html>
<head>
<link href="../<%=c_estilo %>" rel="StyleSheet" type="text/css">
<meta http-equiv="Content-Type" content="text/html; charset=iso-8859-1">
<title>Reportes - Gesti&oacute;n de Desempe&ntilde;o<%if ccodelco<>-1 then%>- RHPro &reg;<%end if%></title>
</head>
<script src="/serviciolocal/shared/js/fn_windows.js"></script>
<script src="/serviciolocal/shared/js/fn_ayuda.js"></script>
<script src="/serviciolocal/shared/js/fn_ay_generica.js"></script>
<script src="/serviciolocal/shared/js/fn_hora.js"></script>
<script>
var titulo='';
function imprimir(){
	parent.frames.ifrm.focus();
	window.print();
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
	titulo='Para las estructuras ';
	titulo= titulo + document.clasifica.datos.nivel1.value;
	l_join= " INNER JOIN his_estructura estr1  ON  (evacab.empleado = estr1.ternro and estr1.estrnro="
			+document.clasifica.datos.estrnro1.value + ") ";
	if (document.clasifica.datos.estrnro2.value!=0){
		titulo= titulo + ', ' + document.clasifica.datos.nivel2.value;
		l_join= l_join + " INNER JOIN his_estructura estr2 ON  (evacab.empleado = estr2.ternro and estr2.estrnro="
				+document.clasifica.datos.estrnro2.value + ") ";
		if (document.clasifica.datos.estrnro3.value!=0){
			titulo= titulo + ', ' + document.clasifica.datos.nivel2.value;
			l_join= l_join + " INNER JOIN his_estructura estr3 ON  (evacab.empleado = estr3.ternro and estr3.estrnro="
					+document.clasifica.datos.estrnro3.value + ") ";
		}
	}
}
return l_join;
}

function cargarlistemp(){
document.datos.listemp.value="";
var l_join='';
document.datos.join.value="";

if (document.datos.quien.value=="estructura"){
	if (document.clasifica.datos.check.checked){
		<%if ccodelco=-1 then%>
		titulo="Todos los Supervisados";
		<%else%>
		titulo="Todos los Evaluados";
		<%end if%>
		document.datos.listemp.value= Nuevo_Dialogo(window, 'rep_resul_totales_eva_02.asp?evaevenro='+document.datos.evaevenro.value, 500,500);
		}
	else
		if (validarestructura()){ 
			var l_join = armarjoin();
			document.datos.join.value=l_join;
			document.datos.listemp.value= Nuevo_Dialogo(window, 'rep_resul_totales_eva_05.asp?evaevenro='+document.datos.evaevenro.value+'&join='+ l_join, 20,10);
		}	
}
else
if (document.datos.quien.value=="evaluador"){
	if ((document.clasifica.datos.ternro.value!="")&&(document.datos.evaevenro.value!="0")){
		<%if ccodelco=-1 then%>
		titulo= "Para el Supervisor "+document.clasifica.datos.empleg.value+'  '+document.clasifica.datos.empleado.value;
		<%else%>
		titulo= "Para el Evaluador "+document.clasifica.datos.empleg.value+'  '+document.clasifica.datos.empleado.value;
		<%end if%>
		document.datos.ternro.value= document.clasifica.datos.ternro.value;	
		document.datos.listemp.value= Nuevo_Dialogo(window, 'rep_resul_totales_eva_04.asp?evaevenro='+document.datos.evaevenro.value+'&evaluador='+ document.clasifica.datos.ternro.value, 20,10);
	}	
}
else{
	if (document.clasifica.datos.ternro.value!=""){
		<%if ccodelco=-1 then%>
			titulo= "Para el Supervisado "+document.clasifica.datos.empleg.value+'  '+document.clasifica.datos.empleado.value;
		<%else%>
			titulo= "Para el Evaluado "+document.clasifica.datos.empleg.value+'  '+document.clasifica.datos.empleado.value;
		<%end if%>
		document.datos.ternro.value= document.clasifica.datos.ternro.value;	
		document.datos.listemp.value= document.clasifica.datos.ternro.value;	
	}	
}		
}

function Validar_Formulario(){
if ((document.datos.evaevenro.value=="0") || (document.datos.evaevenro.value==""))
	alert("Debe seleccionar un Evento.");
else{
cargarlistemp();
if (document.datos.listemp.value=="")
{
	document.ifrm.location="vacio.asp";
	<%if ccodelco=-1 then%>
	alert("No hay Supervisados para esta selección.");
	<%else%>
	alert("No hay empleados para esta selección.");
	<%end if%>
}	
else{
	document.ifrm.location="vacio.asp";
	document.ifrm.location= "rep_emp_formulario_eva_01.asp?evaevenro="+document.datos.evaevenro.value+"&titulo="+ titulo+"&ternro="+document.datos.ternro.value+"&join="+document.datos.join.value+"&quien="+document.datos.quien.value+"&logeadoternro=<%=l_logeadoternro%>&llamadora=<%=l_llamadora%>";
	}
}	
}

function Formulario_Vacio(){
if ((document.datos.evaevenro.value=="0") || (document.datos.evaevenro.value==""))
	alert("Debe seleccionar un Evento.");
else{
cargarlistemp();
if (document.datos.listemp.value=="")
{
	document.ifrm.location="vacio.asp";
	<%if ccodelco=-1 then%>
	alert("No hay Supervisados para esta selección.");
	<%else%>
	alert("No hay empleados para esta selección.");
	<%end if%>
}	
else{
	document.ifrm.location="vacio.asp";
	document.ifrm.location= "rep_vacio_formulario_eva_01.asp?evaevenro="+document.datos.evaevenro.value+"&titulo="+ titulo+"&ternro="+document.datos.ternro.value+"&join="+document.datos.join.value+"&quien="+document.datos.quien.value+"&logeadoternro=<%=l_logeadoternro%>&llamadora=<%=l_llamadora%>";
	}
}	
}

function Archivo(){
if (document.datos.evaevenro.value=="0")
	<% if cint(cdeloitte) = -1 then %>
		alert("Debe seleccionar un Proyecto.");
	<% else%>
		alert("Debe seleccionar un Evento.");
	<% end if%>
else{
cargarlistemp();
if (document.datos.listemp.value=="")
{
	<% if ccodelco=-1 then%>
	alert("No hay Supervisados para esta selección.");
	<% else%>
	alert("No hay empleados para esta selección.");
	<% end if%>
}	
else
	document.ifrm.location= "rep_archivo_formulario_eva_01.asp?evaevenro="+document.datos.evaevenro.value+"&listternro="+document.datos.listemp.value+"&titulo="+ titulo;
}	
}


function buscemp(){
 if (document.datos.quien.value=="empleado")
	window.open('help_emp_01.asp','new','toolbar=no,location=no,directories=no,satus=no,menubar=no,scrollbars=no,resizable=no,width=700,height=400');
}

function Tecla(num){
  if (num==13) {
		abrirVentanaH('nuevo_emp.asp?empleg='+document.datos.empleg.value,'',200,100);
		return false;
  }
  return num;
}

function nuevoempleado(ternro,empleg,terape,ternom)
{
if (empleg != 0) {	
			document.datos.ternro.value = ternro;
			document.datos.empleg.value = empleg;
			document.datos.empleado.value = terape + ", " + ternom;
}
else
{
	<%if ccodelco=-1 then%>
	alert('Supervisado incorrecto');
	<%else%>
	alert('Empleado	incorrecto');
	<%end if%>
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


function Nuevo_Dialogo(w_in, pagina, ancho, alto)
{
 return w_in.showModalDialog(pagina,'', 'center:yes;dialogWidth:' + ancho.toString() + ';dialogHeight:' + alto.toString() + ';');
}
function Ayuda_Fecha(txt)
{
 var jsFecha = Nuevo_Dialogo(window, '/bsas/shared/js/calendar.html', 16, 15);

 if (jsFecha == null) txt.value = ''
 else txt.value = jsFecha;
}


</script>

<body leftmargin="0" rightmargin="0" topmargin="0" bottommargin="0">

<form name="datos"  action="#" method="post">
<input type="Hidden" name="ternro" value="">
<input type="Hidden" name="join" value="">
<input type="Hidden" name="listemp" value="">

<table cellspacing="0" cellpadding="0" border="0" width="100%" >
  <tr>
    <td class="th2" colspan="2">Formulario </td>
	<td nowrap colspan="2" align="right" class="barra" valign="middle">
		<%if cejemplo <> -1 then%>
		<a class=sidebtnABM onclick="Javascript:Formulario_Vacio();" href="#">Reporte Vac&iacute;o</a>
		<%end if%>
		<!--a class=sidebtnABM onclick="Javascript:Archivo();" href="#">Archivo para Merge</a-->
		<a class=sidebtnABM onclick="Javascript:Validar_Formulario();" href="#">Generar Reporte</a>
  		&nbsp;&nbsp;&nbsp;
		<a class=sidebtnSHW href="Javascript:imprimir();">Imprimir</a>
  		&nbsp;&nbsp;&nbsp;
  		<%if ccodelco<>-1 then%>
		<a class=sidebtnHLP href="#" onclick="Javascript:ayuda('<%= Request.ServerVariables("SCRIPT_NAME")%>');">Ayuda</a>
		<%end if%>
	</td>
  </tr>
</table>  
<table cellspacing="0" cellpadding="0" border="0" width="100%" height="100%">
<tr height="10%">
    <td>
		  <b>Eventos de Evaluaci&oacute;n:</b>
		  <br>
		 <select name="evaevenro" size="1" <%if UCase(l_llamadora)="AUTO" and ccodelco=-1 and trim(l_evaevenro)<>"" then %>disabled<%end if%>>
			<option value= 0 >  </option>
		<%	Set l_rs = Server.CreateObject("ADODB.RecordSet")
			l_sql = "SELECT DISTINCT evaevento.evaevenro, evaevedesabr FROM evaevento "
			if UCase(l_llamadora)="AUTO" and ccodelco=-1 then
			l_sql = l_sql & " INNER JOIN evacab ON evacab.evaevenro=evaevento.evaevenro AND evacab.cabaprobada=0"
			l_sql = l_sql & " INNER JOIN evadetevldor ON evadetevldor.evacabnro=evacab.evacabnro AND evatevnro = " & cevaluador & " AND evaluador = " & l_logeadoternro 
			end if
			rsOpen l_rs, cn, l_sql, 0
			do until l_rs.eof	 %>	
			<option value= <%= l_rs("evaevenro") %> > 
			<%= l_rs("evaevenro")&": "& l_rs("evaevedesabr") %> </option>
		<%			l_rs.Movenext
			loop
			l_rs.Close          
			set l_rs=nothing%>	
		</select>
		<script>document.datos.evaevenro.value='<%=l_evaevenro%>';</script>
	</td>
	<td>
	<table cellspacing="0" cellpadding="0" border="0" width="100%">
		<tr>
		<td>
			<input type="Hidden" name="quien" value="estructura">
			<input type="radio" name="quien_"  value="estructura"  onclick="Javascript:radioclick(1);" checked>
			<b>Por Estructura</b> 
		</td>
		<%if UCase(l_llamadora)="AUTO" and trim(l_evaevenro)<>"" and ccodelco=-1 then%>
			<tr><td>&nbsp;</td></tr>
		<%else%>
		<tr>
		<td>
			<input type="radio" name="quien_"  value="evaluador"  onclick="Javascript:radioclick(2);">
			<b>Por <%if ccodelco=-1 then%>Supervisor<%else%>Evaluador<%end if%></b> 
		</td>
		</tr>
		<%end if%>
		<tr>
		<td>
			<input type="radio" name="quien_"  value="empleado"  onclick="Javascript:radioclick(3);">
			<b>Por <%if ccodelco=-1 then%>Supervisado<%else%>Supervisor<%end if%></b> 
		</td>
		</tr>
	</table>
	</td>
	<td>
	<table cellspacing="0" cellpadding="0" border="0" width="100%">
		<tr>
		<td>
				<iframe name="clasifica" src="rep_emp_por_eva_00.asp" width="380" height="112" scrolling="No"></iframe> 
		</td>
		</tr>
	</table>
	</td>
</tr>	
<tr height="80%">
    <td align="right" colspan="3">
		<iframe name="ifrm" src="blanc.asp" width="100%" height="100%"></iframe> 
	</td>
</tr>
<tr height="20%">
    <td align="right" colspan="3" height="40">
	</td>
</tr>
</table>
</form>
</body>
</html>

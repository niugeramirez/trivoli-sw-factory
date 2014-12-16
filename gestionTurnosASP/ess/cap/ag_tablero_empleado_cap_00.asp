<% Option Explicit %>
<!--#include virtual="/serviciolocal/shared/inc/sec.inc"-->
<!--#include virtual="/serviciolocal/shared/inc/const.inc"-->
<!--#include virtual="/serviciolocal/shared/db/conn_db.inc"-->
<!--#include virtual="/serviciolocal/shared/inc/fecha.inc"-->
<!--#include virtual="/serviciolocal/shared/inc/antigfec.inc"-->
<!--#include virtual="/serviciolocal/shared/inc/sqls.inc"-->
<!--
-----------------------------------------------------------------------------
Archivo: ag_tablero_empleado_cap_00.asp
Descripcion: Tablero del Empleado
Autor: Raul Chinestra
-----------------------------------------------------------------------------
-->
<% 
'on error goto 0

Dim l_sql
Dim l_rs
Dim rs9
Dim l_terape 
Dim l_ternom 
Dim l_terape2 
Dim l_ternom2 
Dim l_empleg
Dim l_puenro
Dim l_seleccion

Dim l_empfoto
Dim rs1
Dim sql
Dim l_orinro
Dim l_estnro

Dim l_ternro

Dim siguiente
Dim Anterior
Dim l_orden

dim leg
leg = Session("empleg")
if leg = "" then
    response.write "NO SE HA SELECCIONADO UN EMPLEADO<BR>"
	Response.End
end if
Set l_rs  = Server.CreateObject("ADODB.RecordSet")
l_sql = "SELECT ternro FROM empleado WHERE empleado.empleg =" & leg
l_rs.Open l_sql, cn
if l_rs.eof then
    response.write "NO SE HA SELECCIONADO UN EMPLEADO<BR>"
	response.end
else 
  l_ternro = l_rs("ternro")
end if
l_rs.close
l_empleg     = leg

'l_empleg = request.querystring("empleg")
l_seleccion = request.querystring("seleccion")

if l_empleg = "" then
	Set rs9 = Server.CreateObject("ADODB.RecordSet")
	l_sql = "SELECT min(empleg) FROM empleado where empest = -1"
	RS9.Maxrecords = 1
	rsOpen rs9, cn, l_sql, 0
	if not rs9.eof then
		l_empleg = rs9(0)
	end if
end if

%>

<html>
<head>
<link href="../<%= session("estilo")%>" rel="StyleSheet" type="text/css">
<title> Tablero del Empleado - Capacitación - RHPro &reg;</title>
<script src="/serviciolocal/shared/js/fn_windows.js"></script>
<script src="/serviciolocal/shared/js/fn_confirm.js"></script>
<script src="/serviciolocal/shared/js/fn_ayuda.js"></script>
<script src="/serviciolocal/shared/js/menu_def.js"></script>
<script src="/serviciolocal/shared/js/fn_ay_generica.js"></script>
<script src="/serviciolocal/shared/js/fn_buscar_emp.js"></script>
<script>
function Nuevo_Dialogo(w_in, pagina, ancho, alto)
{
 return w_in.showModalDialog(pagina,'', 'center:yes;dialogWidth:' + ancho.toString() + ';dialogHeight:' + alto.toString());
}

function Tecla(num){
  if (num==13) {
		verificacodigo(document.datos.empleg,document.datos.empleado,'empleg','terape, ternom','empleado');
		Sig_Ant(document.datos.empleg.value);
		return false;
  }
  return num;
}

function emplerror(nro){
	alert('empleado error:'+nro);
}

function datos(){
var dat='';
 dat='ternro='+document.datos.ternro.value+"&empleg="+document.datos.empleg.value+'&empleado='+document.datos.empleado.value;
 dat= dat+ '&fechadesde='+document.datos.fecha.value+ '&fechahasta='+document.datos.fecha.value;
 return dat;
}
function nuevoempleado(ternro,empleg,terape,ternom)
{
if (empleg != 0) 
	{
	document.datos.empleg.value = empleg;
	document.datos.empleado.value = terape + ", " + ternom;
	Sig_Ant(document.datos.empleg.value);
	}
}

function ActualizarGap(){
	document.ifrm.location ="gap_modulos_cap_01.asp?ternro=" + document.datos.ternro.value + "&origen=" + document.datos.orinro.value + "&estado=" + document.datos.estnro.value;
}	

function EstInf(){
	abrirVentana('estudios_informales_cap_00.asp?empleg=' + document.datos.empleg.value,'',700,350);
}	

function EstFor(){
	abrirVentana('estudios_formales_cap_00.asp?empleg=' + document.datos.empleg.value,'',700,350);
}	

function GapMod(){
	abrirVentana('gap_modulos_cap_00.asp?empleg=' + document.datos.empleg.value,'',700,350)

}	
function GapComp(){
	abrirVentana('gap_competencias_cap_00.asp?empleg=' + document.datos.empleg.value,'',700,350)
}	

function Espe(){
	abrirVentana('especializaciones_cap_00.asp?empleg=' + document.datos.empleg.value,'',700,350)
}	

function EliGap(){

 if (ifrm.jsSelRow == null) {
 	alert('Debe seleccionar un Gap')
 }
 else {
	 if ((!(ifrm.jsSelRow.cells(0).innerText.slice(0,6) == 'Manual')) || 
        (!(ifrm.jsSelRow.cells(3).innerText.slice(0,9) == 'Pendiente'))) {
	 	alert('Solamente se permiten Eliminar los Gap Manuales en estado Pendiente');
	 } else {
		eliminarRegistro(document.ifrm,'gap_modulos_cap_04.asp?cabnro=' + document.ifrm.datos.cabnro.value)
	}	 
 } 
}	


function llamadaexcel(){ 
	abrirVentana("gap_modulos_cap_excel.asp?ternro=" + document.datos.ternro.value  + "&origen=" + document.datos.orinro.value + "&estado=" + document.datos.estnro.value,'execl',50,50);
}

</script>

</head>
<body leftmargin="0" topmargin="0" rightmargin="0" bottommargin="0">

<%
Dim l_convnro

Set l_rs = Server.CreateObject("ADODB.RecordSet")
l_sql = "SELECT terape, ternom, ternro, empfoto,terape2, ternom2, puenro"
l_sql = l_sql & " FROM empleado "
l_sql = l_sql & " WHERE empleg=" & l_empleg
l_rs.Maxrecords = 1
rsOpen l_rs, cn, l_sql, 0
l_terape = l_rs("terape")
l_ternom = l_rs("ternom")
l_terape2 = l_rs("terape2")
l_ternom2 = l_rs("ternom2")
l_ternro = l_rs("ternro")
l_puenro = l_rs("puenro")
l_empfoto = trim(l_rs("empfoto") & " ")
l_rs.Close

%>
<form name="datos" action="" method="post">
<input type="hidden" name="ternro" value="<%= l_ternro %>">
<input type="hidden" name="seleccion" value="<%= l_seleccion %>">
<input type="hidden" name="evenro" value=0>
<%

Dim NombrePuesto

l_sql = " SELECT puenro, puedesc  FROM puesto "
l_sql = l_sql & " WHERE puesto.puenro = " & l_puenro

rsOpen l_rs, cn, l_sql, 0
If Not l_rs.EOF Then
	NombrePuesto = l_rs("puedesc")
	l_puenro     = l_rs("puenro")
else 	
	NombrePuesto = "Sin Puesto"
	l_puenro     = 0
End If
l_rs.close

dim salir 

%>
<script>
function Llamar(nro) { 
//alert(document.all.opc2.value);
	switch (parseInt(nro)) {
	   case 1 :
			document.ifrm.location.href="ag_tablero_eventos_cap_01.asp";
			desactive();
			break;
	   case 2 :
		    document.ifrm.location.href="../adp/ag_emp_est_formales_adp_01.asp";
			desactive();
			break;		
	   case 3 :
			document.ifrm.location.href="ag_estudios_informales_cap_01.asp";
			desactive();
			break;		
	   case 4 :
	   		document.ifrm.location.href=src="ag_gap_modulos_cap_01.asp?origen=0&estado=-1";
			desactive();
			break;		
	   case 5 :
	   		//document.ifrm.location.href="estudios_informales_cap_01.asp?";
			//desactive();
			break;		   
	  case 6 :
	   		document.ifrm.location.href="ag_especializaciones_cap_01.asp";
			desactive();
			break;
	  case 7 :
	   		document.ifrm.location.href="ag_gap_competencias_cap_01.asp?ternro=<%= l_ternro %>&estado=2";
			desactive();
			break;
	} 
	document.datos.opc2.value = nro;
}

function active(){
//	document.all.info.className = "sidebtnABM";
//	document.all.info.href = "Javascript:seleccion();";
}

function desactive(){
//	document.all.info.className = "sidebtnDSB";
//	document.all.info.href = "#";
//	document.datos.evenro.value = 0;
}

function seleccion(){
	//alert(document.datos.evenro.value);
	abrirVentana('ag_info_evento_cap_00.asp?evenro='+ document.datos.evenro.value,'',700,550)
}
</script>
<table border="0" cellpadding="0" cellspacing="0"  height="100%" width="100%">
  <tr style="border-color :CadetBlue;">
	<td colspan="3">
		<table border="0" cellpadding="0" cellspacing="0" width="100%" height="0">
			<tr>
	        	<td align="left" class="th2">
					Tablero del empleado
					<!--<a class=sidebtnSHW href="Javascript:window.close();">Salir</a>-->
				</td>
        		<td colspan="2" align="right" class="th2" valign="middle">
                <% 'call MostrarBoton ("sidebtnABM", "Javascript:EstFor();","Estudios Formales")%>&nbsp;
                <% 'call MostrarBoton ("sidebtnABM", "Javascript:EstInf();","Estudios Informales")%>&nbsp;
				<% 'call MostrarBoton ("sidebtnABM", "Javascript:Espe();","Especializaciones")%>&nbsp;
				<% 'call MostrarBoton ("sidebtnABM", "Javascript:GapMod();","Gap Módulos")%>&nbsp;
				<% 'call MostrarBoton ("sidebtnABM", "Javascript:GapComp();","Gap Competencias")%>&nbsp;
				&nbsp;&nbsp;&nbsp;
                <% 'call MostrarBoton ("sidebtnSHW", "Javascript:abrirVentana('help_emp_01.asp?empleado=empleado','',600,400);","Buscar")%>
		        &nbsp;&nbsp;										
				<!--<a class=sidebtnHLP href="Javascript:ayuda('<%'= Request.ServerVariables("SCRIPT_NAME")%>');">Ayuda</a>-->
				</td>
			</tr>
		</table>
	</td>
</tr>
<tr>
	<td align="center">
		<table  border="0" cellpadding="0" cellspacing="0" >
			<tr>
				<td width="50%">&nbsp;</td>
			    <td nowrap align="right"><b>Empleado:</b></td>
				<td colspan="5">
				<input type="text" class="deshabinp" readonly="" value="<%= l_empleg %>" size="8" name="empleg">
				&nbsp;
				<input style="background : #e0e0de;" readonly type="text" name="empleado" size="58" maxlength="45" value="<%= l_terape&" "&l_terape2 &", "&l_ternom&" "&l_ternom2 %>">
				</td>
				<td width="50%">&nbsp;</td>
			</tr>
			<tr>
				<td width="50%">&nbsp;</td>
			    <td nowrap align="right" ><b>Puesto:</b></td>
				<td colspan="5" width="2"><input style="background : #e0e0de;" type="text" name="convenio" size="72" maxlength="50" value='<%= NombrePuesto %>' readonly></td>
				<td width="50%">&nbsp;</td>
			</tr>
			<tr>
				<td width="50%">&nbsp;</td>
				<td nowrap align="right"><b>Origen:</b></td>
				<td colspan="4" align="left" width="1">
					<select name="opc2" onchange="Javascript:Llamar(document.all.opc2.value);">
						<option value=1 selected>Eventos
						<option value=2>Estudios Formales	
						<option value=3>Estudios Informales
						<option value=6>Especializaciones
						<option value=4>Gap por Módulos
						<!--<option value=5>Estudios Informales-->
						<option value=7>Gap Competencias
					</select>
					<script>
						<% If l_seleccion <> "" then %>
							document.datos.opc2.value = <%= l_seleccion %>
						<% End If %>
					</script>
				</td>
				<td width="1">
					&nbsp;
					<!--
					<a class="sidebtnDSB" id="info" href="#">Información del evento</a>
					-->
				</td>
				<td width="50%">&nbsp;</td>
			</tr>
		</table>
	</td>
	</tr>
<tr valign="top" height="100%">
		<td colspan="3" style="">
			<iframe frameborder="0"  scrolling="Yes" name="ifrm" src="ag_tablero_eventos_cap_01.asp" width="100%" height="100%"></iframe> 
		</td>
</tr>
</table>

<% 
cn.Close
set cn = Nothing
%>
</form>	
<SCRIPT SRC="/serviciolocal/shared/js/menu_op.js"></SCRIPT>
</body>
</html>

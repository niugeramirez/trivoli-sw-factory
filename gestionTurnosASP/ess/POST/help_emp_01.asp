<% Option Explicit %>
<!--#include virtual="/ess/ess/shared/inc/const.inc"-->
<!--#include virtual="/serviciolocal/shared/inc/sec.inc"-->
<!--#include virtual="/serviciolocal/shared/inc/const.inc"-->
<!--#include virtual="/serviciolocal/shared/db/conn_db.inc"-->
<!--
Archivo: help_emp_01.asp
Descripción: Usado para encontrar (filtrar) un empleado en datos_ganancias_liq_00.asp
Autor : <Nombre del Creador>
Fecha: <Fecha de Creación>
Modificado:
	Fernando Favre - 31-07-03 - Se agrego la opcion de filtrar los empleados activos, inactivos o todos
	JMH - 22-11-2004 - Se valida que el nombre y apellido sea nombres correctos.
	Martin Ferrro - 16/12/2005 - Se agrego el select de estado
-->
<% 
Dim l_empest
Dim l_rs
Dim l_sql
' Variables
  Dim l_Etiquetas  ' Son los nombres que deben aparecer en la ventana para que el usuario seleccione
  Dim l_Campos     ' Son los campos de la base que apareceran en la clausula where, que deben estar asociados a las etiquetas
  Dim l_Tipos      ' Son los tipos de datos que tienen los campos (N=Numerico, T=Texto y F=Fecha)

' Orden
  Dim l_Orden      ' Son las etiquetas que aparecen en el orden
  Dim l_CamposOr   ' Son los campos para el orden
  
' Orden
  l_Orden     = "Empleado:;Apellido:;Nombre:;Sigla:"
  l_CamposOr  = "empleg;terape;ternom;tidsigla"

l_empest   = request.querystring("empest")
%>
<html>
<head>
<link href="../<%=c_estilo %>" rel="StyleSheet" type="text/css">
<title>RHPro &reg;</title>
<script src="/serviciolocal/shared/js/fn_windows.js"></script>
<script src="/serviciolocal/shared/js/fn_confirm.js"></script>
<script src="/serviciolocal/shared/js/fn_ayuda.js"></script>
<script src="/serviciolocal/shared/js/fn_valida.js"></script>
<script>

function seleccionar(){
ternro=document.ifrm.datos.cabnro.value;
empleg=document.ifrm.datos.empleg.value;
terape=document.ifrm.datos.terape.value;
ternom=document.ifrm.datos.ternom.value;
window.opener.nuevoempleado(ternro,empleg,terape,ternom);
window.close();
}

function buscar(){
filtro="";
and= "";
var correr = true

if (!nombreValido(document.datos.apellido.value)){
	alert("El Apellido contiene caracteres no válidos.")
	correr = false
	}
else
if (!nombreValido(document.datos.nombre.value)){
	alert("El Nombre contiene caracteres no válidos.")
	correr = false
	}

if (document.datos.empest.value != ''){
	filtro= filtro + "empleado.empest = " + document.datos.empest.value;
	and= " and ";
}
if (document.datos.selectape.value == 2){
	filtro= filtro + and + "empleado.terape = '"+document.datos.apellido.value+"'";
	and= " and ";
}
else 
	if (document.datos.selectape.value == 3){
	filtro= filtro + and + "empleado.terape like '"+document.datos.apellido.value+escape("%")+"'";
	and= " and ";
	} 

if (document.datos.selectnom.value == 2){
	filtro= filtro + and + "empleado.ternom = '"+document.datos.nombre.value+"'";
	and= " and ";
}
else 
	if (document.datos.selectnom.value == 3){
	filtro= filtro + and + "empleado.ternom like '"+document.datos.nombre.value+escape("%")+"'";
	and= " and ";
	} 

if ((document.datos.selectleg.value == 2)||(document.datos.selectleg.value == 4)){
	filtro= filtro + and + "empleado.empleg >= "+document.datos.legdesde.value;
	and= " and ";
}
if ((document.datos.selectleg.value == 3)||(document.datos.selectleg.value == 4)){
	filtro= filtro + and +"empleado.empleg <= "+document.datos.leghasta.value;
	and= " and ";
}

if (document.datos.selectdoc.value == 2){
	if (document.datos.doctipo.value != ""){
		filtro= filtro + and + "tipodocu.tidnro = "+document.datos.doctipo.value;
		and= " and ";
	} 
	if (document.datos.selectdocnro.value == 2){
		filtro= filtro + and + "ter_doc.nrodoc = '"+document.datos.docnro.value+"'";
		and= " and ";
	}
	else 
		if (document.datos.selectdocnro.value == 3){
			filtro= filtro + and + "ter_doc.nrodoc like '"+document.datos.docnro.value+escape("%")+"'";
			and= " and ";
		}
}
if (correr)
   //document.ifrm.location="help_emp_02.asp?filtro="+filtro;
   document.ifrm.location="help_emp_02.asp?filtro="+filtro+"&estado="+document.datos.empest.value+"&selectdoc="+document.datos.selectdoc.value;

}


function habdesh(valor,patron,obj){
	if (valor==patron){
		obj.disabled = true;
		obj.value = '';
		obj.style.background= "#e0e0de";
		}
	else {
		obj.disabled = false;
		obj.style.background= "#ffffff";
		}
}

function selectlegajo(){
if (document.datos.selectleg.value == 1){
	habdesh(1,1,document.datos.legdesde);
	habdesh(1,1,document.datos.leghasta);
}
else 
	if (document.datos.selectleg.value == 2){
		habdesh(2,1,document.datos.legdesde);
		habdesh(1,1,document.datos.leghasta);
	}
   	else
		if (document.datos.selectleg.value == 3){
			habdesh(1,1,document.datos.legdesde);
			habdesh(2,1,document.datos.leghasta);
		}
		else {
			habdesh(2,1,document.datos.legdesde);
			habdesh(2,1,document.datos.leghasta);
		}		

}

function nuevabusqueda(){
document.datos.selectape.value=1;
habdesh(1,1,document.datos.apellido);
document.datos.selectnom.value=1;
habdesh(1,1,document.datos.nombre);
document.datos.selectleg.value=1;
habdesh(1,1,document.datos.legdesde);
habdesh(1,1,document.datos.leghasta);
document.datos.selectdoc.value=1;
habdesh(1,1,document.datos.doctipo);
habdesh(1,1,document.datos.selectdocnro);
habdesh(1,1,document.datos.docnro);
document.datos.total.value=0;
document.ifrm.location="help_emp_02.asp?filtro=empleado.ternro=0";
}


function orden(pag)
{
  abrirVentana(encodeURI('orden_param_eyp_00.asp?pagina='+pag+'&lista=<%= l_orden %>&campos=<%= l_camposOr%>&filtro='+document.ifrm.datos.filtro.value),'',350,160)
}

function param(){
	    return ('estado='+document.datos.empest.value + '&selectdoc='+document.datos.selectdoc.value);
}

</script>

</head>
<body leftmargin="0" topmargin="0" rightmargin="0" bottommargin="0" onloand>
<form name="datos" action="" method="post">
<!--input type="Hidden" name="empest" value="<%'=l_empest%>"-->
<table border="0" cellpadding="0" cellspacing="0">
  <tr style="border-color :CadetBlue;">
	<td colspan="2">
		<table border="0" cellpadding="0" cellspacing="0">
			<tr>
	        	<td align="left" class="barra">Busqueda de Empleados</td>
        		<td colspan="2" align="right" class="barra" valign="middle">
				&nbsp;&nbsp;&nbsp;
				<a class=sidebtnSHW href="Javascript:orden('help_emp_02.asp');">Ordenar</a>
				&nbsp;&nbsp;&nbsp;
				<a class=sidebtnHLP href="Javascript:ayuda('<%= Request.ServerVariables("SCRIPT_NAME")%>');">Ayuda</a>
				</td>
			</tr>
		</table>
	</td>
</tr>
<tr>
	<td>
		<table width="100%">
			<tr>
			    <td nowrap align="right"><b>Apellido:</b></td>
				<td>
					<select name="selectape" size="1" onchange="javascript:habdesh(document.datos.selectape.value,1,document.datos.apellido);">
						<option value= 1 > Sin Definir </option>
						<option value= 2 > Igual a </option>
						<option value= 3 > Comienza Con </option>
					</select>

					<input type="text" name="apellido" size="30" maxlength="30" value="" style="background : #e0e0de;" disabled>
				</td>
			</tr>
			<tr>
			    <td nowrap align="right"><b>Nombre:</b></td>
				<td>
					<select name="selectnom" size="1" onchange="javascript:habdesh(document.datos.selectnom.value,1,document.datos.nombre);">
						<option value= 1 > Sin Definir </option>
						<option value= 2 > Igual a </option>
						<option value= 3 > Comienza Con </option>
					</select>

					<input style="background : #e0e0de;" disabled type="text" name="nombre" size="30" maxlength="30" value="">
				</td>
			</tr>
			<tr>
			    <td nowrap align="right"><b>Empleado:</b></td>
				<td>
					<select name="selectleg" size="1" onchange="javascript:selectlegajo();">
						<option value= 1 > Sin Definir  </option>
						<option value= 2 > Desde </option>
						<option value= 3 > Hasta </option>
						<option value= 4 > Entre </option>
					</select>
					&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;
					<input align="right" style="background : #e0e0de;" disabled type="text" name="legdesde" size="11" maxlength="11" value="">
					<b>-</b>
					<input align="right" style="background : #e0e0de;" disabled type="text" name="leghasta" size="11" maxlength="11" value="">

				</td>
			</tr>
			<tr>
			    <td nowrap align="right"><b>Estado:</b></td>
				<td>
					<select name="empest" size="1">
						<option value= "" > Todos  </option>
						<option value= "-1" > Activos </option>
						<option value= "0" > Inactivos </option>
					</select>
					<script>
						document.datos.empest.value = '<%= l_empest %>';
						if ((document.datos.empest.value == '0')||(document.datos.empest.value == '-1')){
							document.datos.empest.disabled = true;
						}
					</script>
				</td>
			</tr>
			<tr>
			    <td nowrap align="right"></td>
				<td>
					<b>Documento:</b>				
				</td>
			</tr>
			<tr>
			    <td nowrap align="right"><b>Tipo:</b></td>
				<td>
					<select name="selectdoc" size="1" 
					onchange="javascript:habdesh(document.datos.selectdoc.value,1,document.datos.doctipo);habdesh(document.datos.selectdoc.value,1,document.datos.selectdocnro);habdesh(1,1,document.datos.docnro);">
						<option value= 1 > Sin Definir </option>
						<option value= 2 > Igual a </option>
					</select>
					<select name="doctipo" size="1"  style="background : #e0e0de;" disabled>
				<%	Set l_rs = Server.CreateObject("ADODB.RecordSet")
					l_sql = "SELECT tidnro, tidsigla, tidnom "
					l_sql  = l_sql  & "FROM tipodocu"
					rsOpen l_rs, Cn, l_sql, 0
					
					do until l_rs.eof	 %>	
						<option value= <%= l_rs("tidnro") %> > 
						<%=  l_rs("tidsigla")&" : "&l_rs("tidnom") %> </option>
					<%			l_rs.Movenext
						loop
						l_rs.Close 		%>	
					</select>
				</td>
			</tr>
			<tr>
			    <td nowrap align="right"><b>Numero:</b></td>
				<td>
					<select name="selectdocnro" size="1"  style="background : #e0e0de;" disabled onchange="javascript:habdesh(document.datos.selectdocnro.value,1,document.datos.docnro);">
						<option value= 1 > Sin Definir </option>
						<option value= 2 > Igual a </option>
						<option value= 3 > Comienza Con </option>
					</select>

					<input  style="background : #e0e0de;" disabled type="text" name="docnro" size="30" maxlength="30" value="">
				</td>
			</tr>
		</table>
	</td>
	<td width="100">
		<table>
			<tr>
				<td>
					<a class=sidebtnABM href="javascript:buscar();" style="HEIGHT:25;vertical-align:middle;">&nbsp;&nbsp;&nbsp;&nbsp;Buscar Ahora&nbsp;&nbsp;&nbsp;&nbsp;</a>
				</td>
			</tr>
			<tr>
				<td>
					<a class=sidebtnABM href="javascript:nuevabusqueda();" style="HEIGHT:25;vertical-align:middle;">&nbsp;Nueva Busqueda&nbsp;</a>
				</td>
			</tr>
			<tr>
				<td>
					<b>Total:</b>
					<input readonly type="text" name="total" size="5" maxlength="5" value="0" style="background : #e0e0de;">
				</td>
			</tr>

		</table>
	</td>
</tr>
<tr>
	<td colspan="2" >
  				<iframe name="ifrm" src="help_emp_02.asp?filtro=empleado.ternro=0" width="100%" height="143"></iframe> 
	</td>
</tr>
</table>
<table width="100%" border="0" cellspacing="0" cellpadding="0">
<tr>
    <td align="right" class="th2">
		<a class=sidebtnABM href="Javascript:seleccionar()">Aceptar</a>
		<a class=sidebtnABM href="Javascript:window.close()">Cancelar</a>
	</td>
</tr>
</table>
</form>	
</body>
</html>

<% Option Explicit %>
<!--#include virtual="/serviciolocal/ess/shared/inc/const.inc"-->
<!--#include virtual="/serviciolocal/shared/inc/sec.inc"-->
<!--#include virtual="/serviciolocal/shared/inc/const.inc"-->
<!--#include virtual="/serviciolocal/shared/db/conn_db.inc"-->
<!--#include virtual="/serviciolocal/shared/inc/fecha.inc"-->
<!--#include virtual="/serviciolocal/shared/inc/antigfec.inc"-->
<!--#include virtual="/serviciolocal/shared/inc/sqls.inc"-->
<!--#include virtual="/serviciolocal/ess/shared/inc/accesoESS.inc"-->
<!--
-----------------------------------------------------------------------------
Archivo		: con_eventos_en_competencia_cap_00.asp
Descripcion	: Consulta de Eventos por competencia
Autor		: Juan Manuel Hoffman
Fecha		: 19/03/2004
-----------------------------------------------------------------------------
-->
<% 
on error goto 0

' Son las listas de parametros a pasarle a los programas de filtro y orden
' En las mismas se deberan poner los valores, separados por un punto y coma

' Filtro
  Dim l_Etiquetas  ' Son los nombres que deben aparecer en la ventana para que el usuario seleccione
  Dim l_Campos     ' Son los campos de la base que apareceran en la clausula where, que deben estar asociados a las etiquetas
  Dim l_Tipos      ' Son los tipos de datos que tienen los campos (N=Numerico, T=Texto y F=Fecha)

' Orden
  Dim l_Orden      ' Son las etiquetas que aparecen en el orden
  Dim l_CamposOr   ' Son los campos para el orden

' Filtro
  l_etiquetas = "C&oacute;d. Ext.:;Descripci&oacute;n:;Curso:"
  l_Campos    = "evecodext;evedesabr;curdesabr"
  l_Tipos     = "T;T;T"

' Orden
  l_Orden     = "C&oacute;d. Ext.:;Descripci&oacute;n:;Curso:"
  l_CamposOr  = "evecodext;evedesabr;curdesabr"

Dim l_orden2
Dim l_sql
Dim l_rs
Dim rs9
Dim l_terape 
Dim l_ternom 
Dim l_terape2 
Dim l_ternom2 
Dim l_empleg
Dim l_compe
Dim l_competencia
Dim l_puenro
Dim l_convnro

Dim l_empfoto
Dim rs1
Dim sql

Dim l_ternro

Dim siguiente
Dim Anterior

l_ternro = l_ess_ternro
l_empleg = l_ess_empleg

l_compe  = request.querystring("competencia")

l_orden2 = request.querystring("orden")
if l_orden2 = "" then
	l_orden2 = "empleg"
end if

' Se busca al empleado
Set l_rs = Server.CreateObject("ADODB.RecordSet")

l_sql = "SELECT terape, ternom, terape2, ternom2, empleg"
l_sql = l_sql & " FROM empleado "
l_sql = l_sql & " WHERE ternro=" & l_ternro

rsOpen l_rs, cn, l_sql, 0

l_terape = l_rs("terape")
l_ternom = l_rs("ternom")
l_terape2 = l_rs("terape2")
l_ternom2 = l_rs("ternom2")
l_empleg  = l_rs("empleg")

l_rs.Close

l_sql = "SELECT evafacnro, evafacdesabr "
l_sql = l_sql & " FROM cap_falencia "
l_sql = l_sql & " INNER JOIN evafactor ON evafactor.evafacnro = cap_falencia.modnro "
l_sql = l_sql & " AND falorigen = 7 "
l_sql = l_sql & " WHERE evafacnro = " & l_compe 

rsOpen l_rs, cn, l_sql, 0

if not l_rs.eof then
   l_competencia = l_rs("evafacnro") & " - " & l_rs("evafacdesabr")
end if   

%>

<html>
<head>
<link href="../<%= c_estilo %>" rel="StyleSheet" type="text/css">
<title> Eventos asociados a la competencia - Capacitación - RHPro &reg;</title>
<script src="/serviciolocal/shared/js/fn_windows.js"></script>
<script src="/serviciolocal/shared/js/fn_confirm.js"></script>
<script src="/serviciolocal/shared/js/fn_ayuda.js"></script>
<script src="/serviciolocal/shared/js/menu_def.js"></script>
<script src="/serviciolocal/shared/js/fn_ay_generica.js"></script>
<script src="/serviciolocal/shared/js/fn_buscar_emp.js"></script>
<script>

function param()
{
	var setear = "ternro=<%= l_ternro %>&competencia= <%= l_compe %>";
	return setear;
}

function orden(pag)
{
  abrirVentana('../shared/asp/orden_param_adp_00.asp?pagina='+pag+'&lista=<%= l_orden %>&campos=<%= l_camposOr%>&filtro='+escape(document.ifrm.datos.filtro.value),'',350,160)
}

function filtro(pag)
{
  abrirVentana('../shared/asp/filtro_param_adp_00.asp?pagina='+pag+'&campos=<%= l_campos%>&tipos=<%=l_tipos%>&etiquetas=<%=l_etiquetas%>&orden='+document.ifrm.datos.orden.value,'',250,160);
}

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


function llamadaexcel(){ 
	abrirVentana("ag_con_eventos_en_competencia_cap_excel.asp?ternro=<%= l_ternro%>&competencia= <%= l_compe %>&orden="+document.ifrm.datos.orden.value+"&filtro="+escape(document.ifrm.datos.filtro.value),'execl',50,50);
}

</script>

</head>
<body leftmargin="0" topmargin="0" rightmargin="0" bottommargin="0" >

<form name="datos" action="" method="post">
<input type="hidden" name="ternro" value="<%= l_ternro %>">
<%

dim salir 

%>

<table border="0" cellpadding="0" cellspacing="0"  height="100%">
  <tr>
	<th colspan="2">
		<table border="0" cellpadding="0" cellspacing="0" width="100%">
			<tr>
	        	<th align="left">
					&nbsp;
				</th>
        		<th colspan="2" align="right" valign="middle">
                    <% 'call MostrarBoton ("sidebtnSHW", "Javascript:llamadaexcel();","Excel")%>
					<a class=sidebtnSHW href="Javascript:llamadaexcel();">Excel</a>
	         	    &nbsp;&nbsp;&nbsp;
	    			<a class=sidebtnSHW href="Javascript:orden('../../cap/ag_con_eventos_en_competencia_cap_01.asp');">Orden</a>
	    			<a class=sidebtnSHW href="Javascript:filtro('../../cap/ag_con_eventos_en_competencia_cap_01.asp');">Filtro</a>
	    			&nbsp;&nbsp;&nbsp;							    			
				</th>
			</tr>
		</table>
	</th>
</tr>
<tr>
	<td>
		<table  border="0" cellpadding="0" cellspacing="0" >
			<tr>
			    <td nowrap align="right"><b>Empleado:</b></td>
				<td colspan="5">
				  <input type="text" style="background : #e0e0de;" readonly value="<%= l_empleg %>" size="8" name="empleg" >&nbsp; &nbsp; &nbsp;
				  <input style="background : #e0e0de;" readonly type="text" name="empleado" size="52" maxlength="45" value="<%= l_terape&" "&l_terape2 &", "&l_ternom&" "&l_ternom2 %>">
				</td>
			</tr>
			<tr>
			    <td nowrap align="right" ><b>Competencia:</b></td>
				<td><input style="background : #e0e0de;" type="text" name="competencia" size="68" maxlength="50" value='<%= l_competencia %>' readonly></td>
			    <td nowrap align="right">&nbsp;</td>
				<td >&nbsp;</td>
			</tr>
		</table>
	</td>
</tr>
<tr>
	<td colspan="2" height="80%">
		<table border="0" cellpadding="0" cellspacing="0"  height="100%">
	        <tr valign="top">
	        	<td align="left" style="">
    	  				<iframe  scrolling="Yes" name="ifrm" src="ag_con_eventos_en_competencia_cap_01.asp?ternro=<%= l_ternro %>&competencia= <%= l_compe %> " width="100%" height="100%"></iframe> 
     				</td>        				
      			</tr>
		</table>
	</td>
</tr>
</table>

<% 
set l_rs = nothing
cn.Close
set cn = Nothing
%>
</form>	
<SCRIPT SRC="/serviciolocal/shared/js/menu_op.js"></SCRIPT>
</body>
</html>

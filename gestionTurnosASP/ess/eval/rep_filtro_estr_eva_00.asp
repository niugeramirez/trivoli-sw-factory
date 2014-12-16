<% Option Explicit %>
<!--#include virtual="/ess/ess/shared/inc/const.inc"-->
<!--#include virtual="/serviciolocal/shared/db/conn_db.inc"-->

<%
'---------------------------------------------------------------------------------
'Archivo	: rep_filtro_estr_eva_00.asp
'Descripción: 
'Autor		: 
'Fecha		: 
'Modificado	: 04-08-2003 - reemplazar llamadas a gengrup_v2_filtroNivel 
'			  por filtroNivel.asp
'Modificado	: 05-08-2003-CCRossi -Sacar select de empleado que ya no se usan
'Modificado: 17-08-2005 CCRossi Agrandar selets de estructuras
'----------------------------------------------------------------------------------

Dim l_max 
Dim l_min 
Dim l_rs 
Dim l_sql 
Dim l_select

l_select = "<option value= 0 selected></option>"

Set l_rs = Server.CreateObject("ADODB.RecordSet")
l_sql = "SELECT tenro, tedabr "
l_sql  = l_sql  & "FROM tipoestructura Order By tedabr"
rsOpen l_rs, cn, l_sql, 0 
do until l_rs.eof 
  l_select = l_select & "<option value=" & l_rs("tenro") & ">" & l_rs("tedabr") & "</option>"
  l_rs.Movenext
loop
l_rs.Close
l_select = l_select & "</select>"
%>

<html>
<head>
<link href="../<%=c_estilo %>" rel="StyleSheet" type="text/css">
<meta http-equiv="Content-Type" content="text/html; charset=iso-8859-1">
<title>Filtrar Estructuras - Gesti&oacute;n de Desempe&ntilde;o - RHPro &reg;</title>
</head>
<script>
function filtrar()
{
	var tex1="";
	var niv1=0;
	var tex2="";
	var niv2=0;
	var tex3="";
	var niv3=0;
	
	if (document.datos.tiponivel1.value != 0)
	  {
  	  if (document.nivel1.registro.grupo.value == -1)
	    {
		alert ('Debe seleccionar algún valor para el filtro 1');
		return;
		}
	  else
	    {
		tex1 =document.datos.tiponivel1.options[document.datos.tiponivel1.selectedIndex].text+ " -> "+document.nivel1.registro.grupo.options[document.nivel1.registro.grupo.selectedIndex].text;
		niv1 = document.nivel1.registro.grupo.value ;
		}
	  }
	if (document.datos.tiponivel2.value != 0)
	  {
  	  if (document.nivel2.registro.grupo.value == -1)
	    {
		alert ('Debe seleccionar algún valor para el filtro 2');
		return;
		}
	  else
	    {
		tex2 =document.datos.tiponivel2.options[document.datos.tiponivel2.selectedIndex].text+ " -> "+document.nivel2.registro.grupo.options[document.nivel2.registro.grupo.selectedIndex].text;
		niv2 = document.nivel2.registro.grupo.value ;
		}
	  }
	if (document.datos.tiponivel3.value != 0)
	  {
  	  if (document.nivel3.registro.grupo.value == -1)
	    {
		alert ('Debe seleccionar algún valor para el filtro 3');
		return;
		}
	  else
	    {
		tex3 =document.datos.tiponivel3.options[document.datos.tiponivel3.selectedIndex].text+ " -> "+document.nivel3.registro.grupo.options[document.nivel3.registro.grupo.selectedIndex].text;
		niv3 = document.nivel3.registro.grupo.value ;
		}
	  }
	opener.ponerestructuras(niv1,tex1,niv2,tex2,niv3,tex3);  
   	close();
	
}

function ajustacontroles()
{

if (document.datos.tiponivel1.value == 0)
  {
  document.datos.tiponivel2.disabled = true;
  document.nivel2.registro.grupo.disabled = true;
  }
else  
  {
  document.datos.tiponivel2.disabled = false;
  document.nivel2.registro.grupo.disabled = false;
  }
if ((document.datos.tiponivel2.value == 0) ||
    (document.datos.tiponivel2.disabled == true))
  {
  document.datos.tiponivel3.disabled = true;
  document.nivel3.registro.grupo.disabled = true;
  }
else  
  {
  document.datos.tiponivel3.disabled = false;
  document.nivel3.registro.grupo.disabled = false;
  }
}

function actualiza(destino, valor)
{
ajustacontroles();
destino.location = 'filtroNivel.asp?tipo=' + valor;
}

</script>
<style>
.iframe{
	width: 405;
	height: 25;
	border: 0px solid White;
}
</style>
<body leftmargin="0" rightmargin="0" topmargin="0" bottommargin="0" scroll=no>
<form name="datos" method="post">
<table cellspacing="1" cellpadding="0" border="0" width="100%">
  <tr>
    <td class="th2" colspan="3">Filtrar Estructuras</td>
  </tr>
<tr>
    <td align="right" rowspan="2"><b>Nivel 1:</b></td>
	<td><b>Tipo:</b></td>
    <td>
	<select name='tiponivel1' size='1'  onchange="javascript:actualiza(document.nivel1, datos.tiponivel1.value)">           
    <%= l_select %>		 
	</td>
</tr>
<tr>
	<td><b>Valor:</b></td>
    <td>
    <iframe class="iframe"  name="nivel1" scrolling="No" src="filtroNivel.asp?tipo=0" width="405" height="25"></iframe> 
	</td>
</tr>
<tr>
    <td align="right" rowspan="2"><b>Nivel 2:</b></td>
	<td><b>Tipo:</b></td>
    <td>
	<select disabled name='tiponivel2' size='1' onchange="javascript:actualiza(document.nivel2, datos.tiponivel2.value)">	
    <%= l_select %>		 
	</td>
</tr>
<tr>
	<td><b>Valor:</b></td>
    <td>
    <iframe class="iframe" name="nivel2" scrolling="No" src="filtroNivel.asp?tipo=0&disabled=disabled" width="405" height="25"></iframe> 
	</td>
</tr>
<tr>
    <td align="right" rowspan="2"><b>Nivel 3:</b></td>
	<td><b>Tipo:</b></td>
    <td>
	<select disabled name='tiponivel3' size='1' onchange="javascript:actualiza(document.nivel3, datos.tiponivel3.value)">	
    <%= l_select %>		 
	</td>
</tr>
<tr>
	<td><b>Valor:</b></td>
    <td>
    <iframe class="iframe" name="nivel3" scrolling="No" src="filtroNivel.asp?tipo=0&disabled=disabled" width="405" height="25"></iframe> 
	</td>
</tr>

</table>
<table width="100%" border="0" cellspacing="0" cellpadding="0">
<tr>
    <td align="right" class="th2">
		<a class=sidebtnABM href="Javascript:filtrar()">Aceptar</a>
		<a class=sidebtnABM href="Javascript:window.close()">Cancelar</a>		
	</td>
</tr>
</table>
</form>
</body>
</html>

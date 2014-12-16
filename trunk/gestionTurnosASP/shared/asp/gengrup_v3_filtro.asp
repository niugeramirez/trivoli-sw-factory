<% Option Explicit %>
<!--#include virtual="/turnos/shared/db/conn_db.inc"-->
<!--#include virtual="/turnos/shared/inc/sqls.inc"-->

<%
Dim l_max 
Dim l_min 
Dim l_rs 
Dim l_sql 
Dim l_select

Set l_rs = Server.CreateObject("ADODB.RecordSet")

l_sql = "SELECT empleg FROM v_empleado ORDER BY empleg desc "
fsql_first l_sql,1
rsOpen l_rs, cn, l_sql, 0 

l_max = l_rs(0)
l_rs.Close

l_sql = "SELECT empleg FROM v_empleado ORDER BY empleg "
fsql_first l_sql,1
rsOpen l_rs, cn, l_sql, 0 

l_min = l_rs(0)
l_rs.Close

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
<link href="/turnos/shared/css/tables3.css" rel="StyleSheet" type="text/css">
<meta http-equiv="Content-Type" content="text/html; charset=iso-8859-1">
<title><%= Session("Titulo")%>Filtrar - Gesti&oacute;n de Tiempos - RHPro &reg;</title>
</head>
<script>
function filtrar()
{
    var tex;
    var cant;
	var tex2;
	tex = "";
	cant = 0;
	
    if (document.datos.estado[0].checked)
	  tex =  "(empest = -1)"
    if (document.datos.estado[1].checked)
	  tex =  "(empest = 0)"
    if (document.datos.estado[2].checked)
	  tex =  "(0 = 0)"

    if (isNaN(document.datos.legdesde.value) || isNaN(document.datos.leghasta.value))
	  {
	    alert('El legajo debe ser numérico');
		return;
	  }
    tex = tex + " AND (empleg >= " + document.datos.legdesde.value + ") AND (empleg <= " + document.datos.leghasta.value + ")" 
	if (document.datos.tiponivel1.value != 0)
	  {
  	  if (document.nivel1.registro.grupo.value == -1)
	    {
		alert ('Debe seleccionar algún valor para el filtro 1');
		return;
		}
	  else
	    {
		cant = 1;
		tex2 = " (select * from his_estructura where v_empleado.ternro = his_estructura.ternro and his_estructura.htethasta IS NULL and (((his_estructura.tenro = " + document.datos.tiponivel1.value + ") AND ( his_estructura.estrnro = " + document.nivel1.registro.grupo.value + "))";
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
		cant = 2;
		tex2 = tex2 + " OR ((his_estructura.tenro = " + document.datos.tiponivel2.value + ") AND ( his_estructura.estrnro = " + document.nivel2.registro.grupo.value + "))"
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
		cant = 3;
		tex2 = tex2 + " OR ((his_estructura.tenro = " + document.datos.tiponivel3.value + ") AND ( his_estructura.estrnro = " + document.nivel3.registro.grupo.value + "))"
		}
	  }
	if (cant != 0)  
	{
	//tex2 = " " + cant + tex2 + "))";
	tex2 = "EXISTS " + tex2 + "))";
	tex = tex + " AND " + tex2;
//	alert(tex)
	}
    if (document.datos.restaurar.checked)
	  tex =  "(0 = 0)"
    parent.returnValue = tex;
	parent.close();
   	return true;
}

function ajustacontroles()
{

if (document.datos.tiponivel1.value == 0)
  {
  document.datos.tiponivel2.disabled = true;
  //document.nivel2.registro.grupo.disabled = true;
  }
else  
  {
  document.datos.tiponivel2.disabled = false;
  //document.nivel2.registro.grupo.disabled = false;
  }
if ((document.datos.tiponivel2.value == 0) ||
    (document.datos.tiponivel2.disabled == true))
  {
  document.datos.tiponivel3.disabled = true;
  //document.nivel3.registro.grupo.disabled = true;
  }
else  
  {
  document.datos.tiponivel3.disabled = false;
  //document.nivel3.registro.grupo.disabled = false;
  }
}

function actualiza(destino, valor)
{
ajustacontroles();
destino.location = 'gengrup_v3_filtroNivel.asp?tipo=' + valor;
}

parent.dialogHeight = 22;
window.resizeTo(300,180)

</script>

<body leftmargin="0" rightmargin="0" topmargin="0" bottommargin="0" scroll=no>
<form name="datos" method="post">
<table cellspacing="1" cellpadding="0" border="0" width="100%">
  <tr>
    <td class="th2" colspan="3">Filtrar</td>
  </tr>
<tr>
    <td align="right"><b>Desde número:</b></td>
	<td colspan="2"><input type="Text" name="legdesde" value="<%= l_min %>" size="5">&nbsp;<b>hasta</b>&nbsp;
  	    <input type="Text" name="leghasta" value="<%= l_max %>" size="5"></td>
</tr>
<tr>
    <td align="right"><b>Estado:</b></td>
	<td colspan="2"><input type="Radio" name="estado" value="-1" checked> Activo &nbsp;
  	    <input type="Radio" name="estado" value="0"> Inactivo &nbsp;
	    <input type="Radio" name="estado" value="1"> Ambos &nbsp;</td>
</tr>
<tr>
    <td align="right" rowspan="2"><b>Estructura de nivel 1:</b></td>
	<td><b>Tipo:</b></td>
    <td>
	<select name='tiponivel1' size='1'  onchange="javascript:actualiza(document.nivel1, datos.tiponivel1.value)">           
    <%= l_select %>		 
	</td>
</tr>
<tr>
	<td><b>Valor:</b></td>
    <td>
    <iframe name="nivel1" scrolling="No" src="gengrup_v3_filtroNivel.asp?tipo=0" width="155" height="25"></iframe> 
	</td>
</tr>
<tr>
    <td align="right" rowspan="2"><b>Estructura de nivel 2:</b></td>
	<td><b>Tipo:</b></td>
    <td>
	<select disabled name='tiponivel2' size='1' onchange="javascript:actualiza(document.nivel2, datos.tiponivel2.value)">	
    <%= l_select %>		 
	</td>
</tr>
<tr>
	<td><b>Valor:</b></td>
    <td>
    <iframe name="nivel2" scrolling="No" src="gengrup_v3_filtroNivel.asp?tipo=0&disabled=disabled" width="155" height="25"></iframe> 
	</td>
</tr>
<tr>
    <td align="right" rowspan="2"><b>Estructura de nivel 3:</b></td>
	<td><b>Tipo:</b></td>
    <td>
	<select disabled name='tiponivel3' size='1' onchange="javascript:actualiza(document.nivel3, datos.tiponivel3.value)">	
    <%= l_select %>		 
	</td>
</tr>
<tr>
	<td><b>Valor:</b></td>
    <td>
    <iframe name="nivel3" scrolling="No" src="gengrup_v3_filtroNivel.asp?tipo=0&disabled=disabled" width="155" height="25"></iframe> 
	</td>
</tr>


<tr>
    <td align="right"><b>Restaurar:</b></td>
	<td colspan="2"><input type="Checkbox" name="restaurar"></td>
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

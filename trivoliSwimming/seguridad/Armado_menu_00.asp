<% Option Explicit %>
<!--#include virtual="/turnos/shared/inc/sec.inc"-->
<!--#include virtual="/turnos/shared/inc/const.inc"-->
<!--#include virtual="/turnos/shared/db/conn_db.inc"-->
<!--
-----------------------------------------------------------------------------
Archivo        : Armado_menu_00.asp
Descripcion    : 
Modificacion   :
-----------------------------------------------------------------------------
-->
<% 
Dim rs
Dim sql

Set rs = Server.CreateObject("ADODB.RecordSet")
%>
<html>
<head>
<link href="/turnos/shared/css/tables3.css" rel="StyleSheet" type="text/css">
<title><%= Session("Titulo")%>Armado por Men&uacute; - Ticket</title>
<script src="/turnos/shared/js/fn_windows.js"></script>
<script src="/turnos/shared/js/fn_confirm.js"></script>
<script src="/turnos/shared/js/fn_ayuda.js"></script>
<script src="/turnos/shared/js/fn_ay_generica.js"></script>
<script>
function actualizar()
{
if (document.datos.menuraiz.value != "-1")
	document.ifrm.location = "armado_menu_01.asp?menuraiz=" + document.datos.menuraiz.value
else
	document.ifrm.location = "blanc.html";
document.menu2.location = "blanc.html";
}

function moverderecha()
{
  if (ifrm.jsSelRow == null)
    alert("Debe seleccionar un registro.")
  else	
    {
	if (ifrm.jsSelRow.cells(6).innerText < 5)
	  {
      ifrm.jsSelRow.cells(0).innerText = "- " + ifrm.jsSelRow.cells(0).innerText;
	  ifrm.jsSelRow.cells(6).innerText = (ifrm.jsSelRow.cells(6).innerText * 1) + 1;
	  }
	else
	  alert("No puede haber un nivel mayor que 5.");  
	}
}

function moverizquierda()
{
  if (ifrm.jsSelRow == null)
    alert("Debe seleccionar un registro.")
  else	
    {
	if (ifrm.jsSelRow.cells(6).innerText > 0)
	  {
      ifrm.jsSelRow.cells(0).innerText = ifrm.jsSelRow.cells(0).innerText.substr(2);
	  ifrm.jsSelRow.cells(6).innerText = (ifrm.jsSelRow.cells(6).innerText) - 1;
	  }
	else
	  alert("No se puede bajar de nivel.");  
	}
}

function _agregaRegistro()
{
    var mynewRow;
	var desde;
    if (ifrm.jsSelRow == null)
    {
      mynewRow = ifrm.tabla.insertRow(0);
	  desde = 1;
	}
    else
	{
      mynewRow = ifrm.tabla.insertRow(ifrm.jsSelRow.cells(1).innerText);
	  desde = ifrm.jsSelRow.cells(1).innerText;
  	}
	mynewRow.insertCell();
	mynewRow.cells(0).innerText = " ";
	mynewRow.insertCell();
    mynewRow.cells(1).innerText = desde + " ";
//    mynewRow.cells(1).style.display = "none";	
	alert(desde);
    for (i=desde; i < ifrm.tabla.rows.length; i++) 	
      ifrm.tabla.rows(i).cells(1).innerText = ifrm.tabla.rows(i).cells(1).innerText * 1 + 1;	

//  document.ifrm.location = "armado_menu_01.asp?menuraiz=" + document.datos.menuraiz.value + "&orden=" + ifrm.jsSelRow.cells(1).innerText;
}

function altahijo()
{
if (document.ifrm.jsSelRow == null)
	alert("Debe seleccionar un item para asignarle un hijo.")
else
    {
		var donde
		donde = 'armado_menu_05.asp?menuaccess=' + document.menu2.datos.menuaccess.value;
		donde += '&menuname=' + document.menu2.datos.nombre.value;
		donde += '&action=' + escape(document.menu2.datos.accion.value);
		donde += '&menuimg=' + document.menu2.datos.menuimg.value;
		donde += '&menuorder=' + document.ifrm.jsSelRow.cells(2).innerText;
		donde += '&menuraiz=' + document.ifrm.jsSelRow.cells(3).innerText;
		donde += '&tipo=HIJO';
		abrirVentanaH(donde,'',150,150);
    }
}

function altapar()
{
var orden;
var donde;
var padre;

if (document.ifrm.jsSelRow == null)
  {
    orden = ifrm.tabla.rows(0).cells(2).innerText - 1;
	padre = ifrm.tabla.rows(0).cells(4).innerText; 
  }	
else
  {
    orden = document.ifrm.jsSelRow.cells(2).innerText;
    padre = document.ifrm.jsSelRow.cells(4).innerText;
  }	
		donde = 'armado_menu_05.asp?menuaccess=' + document.menu2.datos.menuaccess.value;
		donde += '&menuname=' + document.menu2.datos.nombre.value;
		donde += '&action=' + escape(document.menu2.datos.accion.value);
		donde += '&menuimg=' + document.menu2.datos.menuimg.value;
		donde += '&menuorder=' + orden;
		donde += '&parent=' + padre;
		donde += '&menuraiz=' + document.ifrm.jsSelRow.cells(3).innerText;
		donde += '&tipo=PAR'
		abrirVentanaH(donde,'',150,150);
}

function agregaRegistro()
{
if (document.ifrm.document.datos != null)
  document.menu2.location = "armado_menu_02.asp?tipo=A";
else
  alert("Debe seleccionar una raiz.")  
}

function modificar()
{
if (document.ifrm.jsSelRow == null)
	alert("Debe seleccionar un item.")
else
	{
	if (document.menu2.datos.nombre.value == "")
	  alert("Debe ingresar un nombre.");
	else
		{  
		var donde
		donde = 'armado_menu_03.asp?menuaccess=' + document.menu2.datos.menuaccess.value;
		donde += '&menuname=' + document.menu2.datos.nombre.value;
		donde += '&action=' + escape(document.menu2.datos.accion.value);
		donde += '&menuimg=' + document.menu2.datos.menuimg.value;
		donde += '&menuorder=' + document.menu2.datos.menuorder.value;
		donde += '&menuraiz=' + document.menu2.datos.menuraiz.value;
		donde += '&nivel=' + ifrm.jsSelRow.cells(6).innerText;
		donde += '&parent=' + ifrm.jsSelRow.cells(4).innerText;
		abrirVentanaH(donde,'',150,150);
		}
	}
}

function botones()
{
if (document.ifrm.jsSelRow == null)
	alert("Debe seleccionar un item.")
else
	{
 	var donde
	donde = 'armado_menu_10.asp?menuname=' + document.menu2.datos.nombre.value;
	donde += '&menuorder=' + document.menu2.datos.menuorder.value;
	donde += '&menuraiz=' + document.menu2.datos.menuraiz.value;
	abrirVentana(donde,'',550,250);
	}
}


function eliminaRegistro()
{
if (document.ifrm.jsSelRow == null)
	alert("Debe seleccionar un item.")
else
	{
    if (ifrm.tabla.rows.length != ifrm.jsSelRow.cells(1).innerText)
	  {
	  if (ifrm.jsSelRow.cells(6).innerText < ifrm.tabla.rows(ifrm.jsSelRow.cells(1).innerText * 1).cells(6).innerText)
	    {
  	      alert('El registro tiene otros dependientes.\nDebe eliminarlos primero.'); 	
 	  	  return;
		}
	  }
	var donde
	donde = 'armado_menu_04.asp?menuorder=' + document.menu2.datos.menuorder.value;
	donde += '&menuraiz=' + document.menu2.datos.menuraiz.value;
	donde += '&parent=' + ifrm.jsSelRow.cells(4).innerText;
	abrirVentanaH(donde,'',150,150);
	}
}

</script>

</head>
<body leftmargin="0" topmargin="0" rightmargin="0" bottommargin="0">
<table border="0" cellpadding="0" cellspacing="0" width="100%" height="100%">
	<tr style="border-color :CadetBlue;">
		<td align="left" class="barra">Armado del Men&uacute;</td>
		<td align="right" class="barra">
			<a class=sidebtnHLP href="Javascript:ayuda('<%= Request.ServerVariables("SCRIPT_NAME")%>');">Ayuda</a>
		</td>
	</tr>
<form name="datos">
	<tr>
		<td colspan="2">
			<b>Ra&iacute;z:</b>
			<% 
			sql = "SELECT menunro, menudesc, menudescext FROM menuraiz order by menudesc"
			rsOpen rs, cn, sql, 0
			%>		     
			<select name="menuraiz" onchange="Javascript:actualizar()">
			<option value="-1" SELECTED>Ninguno</option>
			<% do until rs.eof %>
				<option value="<%= rs("menunro") %>"><%= rs("menudescext") %></option>
			<%
			rs.MoveNext
			loop
			rs.Close
			%>
			</select>
		</td>
	</tr>
	<tr>
		<td colspan="2">
			<b>Menu:</b>
		</td>
	</tr>
	<tr>
		<td height="100%">
			<iframe name="ifrm" src="blanc.html" width="100%" height="100%"></iframe> 
		</td>
		<td align="center" class="barra" width="10%">
			<a class=sidebtnABM href="Javascript:agregaRegistro()">Agregar</a><br><br><br>
			<a class=sidebtnABM href="Javascript:eliminaRegistro()">Borrar</a><br><br><br>
		</td>
	</tr>
	<tr>
		<td height="160" colspan="2">
			<b>Datos a Modificar:</b><br>
			<iframe name="menu2" src="blanc.html" width="100%" height="150" scrolling="No"></iframe> 
		</td>
	</tr>
</form>
</table>
</body>
</html>

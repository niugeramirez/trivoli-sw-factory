<%Option Explicit %>
<!--#include virtual="/ess/ess/shared/inc/const.inc"-->
<!--#include virtual="/serviciolocal/shared/inc/sec.inc"-->
<!--#include virtual="/serviciolocal/shared/inc/const.inc"-->
<!--#include virtual="/serviciolocal/shared/db/conn_db.inc"-->
<!--
Archivo    : pedido_emp_vac_gti_00.asp
Descripción: Pedido de vacaciones
Autor      : Scarpa D.
Fecha      : 08/10/2004
Modificado : 
-->
<% 
  on error goto 0
  
' Variables
  Dim l_ternro
  Dim l_empleg
  Dim l_terapenom
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
  l_etiquetas = "Período:;Fecha desde:;Fecha hasta:;Cantidad:;Estado:;Días Hábiles:;Días Feriados:"
  l_Campos    = "vacdesc;vdiapeddesde;vdiapedhasta;vdiapedcant;vdiaspedestado;vdiaspedhabiles;vdiaspedferiados"
  l_Tipos     = "T;F;F;N;N;N;"

' Orden
  l_Orden     = "Período:;Fecha desde:;Fecha hasta:;Cantidad:;Estado:;Días Hábiles:;Días Feriados:"
  l_CamposOr  = "vacdesc;vdiapeddesde;vdiapedhasta;vdiapedcant;vdiaspedestado;vdiaspedhabiles;vdiaspedferiados"

Set l_rs  = Server.CreateObject("ADODB.RecordSet")

dim leg
leg = Session("empleg")
if leg = "" then
    response.write "NO SE HA SELECCIONADO UN EMPLEADO<BR>"
	Response.End
end if

l_sql = "SELECT ternro FROM empleado WHERE empleado.empleg = " & leg
l_rs.Open l_sql, cn
if l_rs.eof then
    response.write "NO SE HA SELECCIONADO UN EMPLEADO<BR>"
	response.end
else 
  l_ternro = l_rs("ternro")
end if
l_rs.close

%>

<html>
<head>
<link href="../<%=c_estilo %>" rel="StyleSheet" type="text/css">
<meta http-equiv="Content-Type" content="text/html; charset=iso-8859-1">
<title>Pedido de Vacaciones</title>
<script src="/serviciolocal/shared/js/fn_windows.js"></script>
<script src="/serviciolocal/shared/js/fn_confirm.js"></script>
<script src="/serviciolocal/shared/js/fn_ayuda.js"></script>
<script>

function orden(pag)
{
  abrirVentana('pedido_emp_vac_gti_100.asp?pagina='+pag+'&lista=<%= l_orden %>&campos=<%= l_camposOr%>&filtro='+escape(document.ifrm.datos.filtro.value),'',350,160)
}

function filtro(pag)
{
  abrirVentana('pedido_emp_vac_gti_99.asp?pagina='+pag+'&campos=<%=l_campos%>&tipos=<%=l_tipos%>&etiquetas=<%=l_etiquetas%>&orden='+document.ifrm.datos.orden.value,'',250,160);	
}
    	
function param(){
	var chequear;
	chequear=  'ternro=' + document.datos.ternro.value;
	return chequear;
}

function modificar(){
    if (document.ifrm.datos.cabnro.value == ''){
	   alert('Debe seleccionar un registro.');
	}else{
	    if (document.ifrm.datos.cabnro.value != '0'){
	       abrirVentanaVerif('pedido_emp_vac_gti_02.asp?Tipo=M&vdiapednro=' + document.ifrm.datos.cabnro.value,'',620,290)
		}else{
		   alert('No se pueden modificar las licencias aceptadas.');
		}
	}
}    	   

function borrar(){
    if (document.ifrm.datos.cabnro.value == ''){
	   alert('Debe seleccionar un registro.');
	}else{
	    if (document.ifrm.datos.cabnro.value != '0'){
	       eliminarRegistro(document.ifrm,'pedido_emp_vac_gti_04.asp?vdiapednro=' + document.ifrm.datos.cabnro.value);
		}else{
		   alert('No se pueden borrar las licencias aceptadas.');
		}
	}
}
</script>

</head>
<body leftmargin="0" topmargin="0" rightmargin="0" bottommargin="0">
<form name=datos>
	<input type="hidden" name="ternro" value="<%=l_ternro%>">
</form>

<table border="0" cellpadding="0" cellspacing="0" height="100%">
<tr style="border-color :CadetBlue;">
	<th align="left">Pedido de Vacaciones</th>
    <th align="right">
		<a class=sidebtnABM href="Javascript:abrirVentana('pedido_emp_vac_gti_02.asp?Tipo=A','',620,290)">Alta</a>
		<a class=sidebtnABM href="Javascript:borrar();">Baja</a>
		<a class=sidebtnABM href="Javascript:modificar();">Modifica</a>
		&nbsp;&nbsp;&nbsp;
		<a class=sidebtnSHW href="Javascript:orden('pedido_emp_vac_gti_01.asp');">Orden</a>
		<a class=sidebtnSHW href="Javascript:filtro('pedido_emp_vac_gti_01.asp');">Filtro</a>
    </th>
</tr>

<tr valign="top">
   <td colspan="2" style="" height="98%">
   <iframe name="ifrm" src="pedido_emp_vac_gti_01.asp?ternro=<%=l_ternro%>" width="100%" height="100%"></iframe> 
   </td>
</tr>

</table>
</body>
</html>

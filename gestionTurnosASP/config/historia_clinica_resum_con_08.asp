<% Option Explicit
if request.querystring("excel") then
	Response.AddHeader "Content-Disposition", "attachment;filename=Planilla de Turnos.xls" 
	Response.ContentType = "application/vnd.ms-excel"
end if
 %>
<!--#include virtual="/turnos/shared/db/conn_db.inc"-->
<!--#include virtual="/turnos/shared/inc/fecha.inc"-->
<% 

Dim l_rs
Dim l_sql
Dim l_filtro
Dim l_orden
Dim l_sqlfiltro
Dim l_sqlorden
Dim l_paciente
Dim l_medico
dim l_detalle

dim l_dni
dim l_fechanacimiento
dim l_domicilio
dim l_telefono 

dim l_fondo

Dim l_id

Dim l_primero

l_filtro = request("filtro")
l_orden  = request("orden")
l_id     = request("id")


%>
<!DOCTYPE HTML PUBLIC "-//IETF//DTD HTML//EN">
<html>
<script src="/turnos/shared/js/fn_windows.js"></script>
<script src="/turnos/shared/js/fn_confirm.js"></script>
<script src="/turnos/shared/js/fn_ayuda.js"></script>

<head>
<% if request.querystring("excel") = false then  %>
<link href="/turnos/ess/shared/css/tables_gray.css" rel="StyleSheet" type="text/css">
<% End If %>
<meta http-equiv="Content-Type" content="text/html; charset=iso-8859-1">
<title>Vista Preliminar Historia Clinica</title>
</head>

<script>
var jsSelRow = null;

function Deseleccionar(fila){
	fila.className = "MouseOutRow";
}

function Seleccionar(fila,cabnro, turnoid){
	if (jsSelRow != null){
		Deseleccionar(jsSelRow);
	};
	document.datos.cabnro.value = cabnro;
	document.datos.idturno.value = turnoid;
	fila.className = "SelectedRow";
	jsSelRow = fila;
}

</script>

<script language="Javascript">
	function imprSelec(nombre) {
	  var ficha = document.getElementById(nombre);
	  var ventimp = window.open(' ', 'popimpr');
	  ventimp.document.write( ficha.innerHTML );
	  ventimp.document.close();
	  ventimp.print( );
	  ventimp.close();
	}
	</script>
<% 

Set l_rs = Server.CreateObject("ADODB.RecordSet")

' Obtengo la cantidad de turnos simultaneos del Recurso Reservable
l_sql = "SELECT  historia_clinica_resumida.* , clientespacientes.apellido, clientespacientes.nombre, recursosreservables.descripcion , clientespacientes.dni, clientespacientes.fechanacimiento , clientespacientes.domicilio, clientespacientes.telefono "
l_sql = l_sql & " FROM historia_clinica_resumida "
l_sql = l_sql & " LEFT JOIN clientespacientes ON clientespacientes.id = historia_clinica_resumida.idclientepaciente "
l_sql = l_sql & " LEFT JOIN recursosreservables ON recursosreservables.id = historia_clinica_resumida.idrecursoreservable "
l_sql = l_sql & " WHERE historia_clinica_resumida.id = " & l_id
rsOpen l_rs, cn, l_sql, 0 
if not l_rs.eof then
	l_paciente = l_rs("apellido") & " " & l_rs("nombre") 
	l_medico = l_rs("descripcion")
	l_detalle = l_rs("detalle")
	l_dni = l_rs("dni")
	l_fechanacimiento = l_rs("fechanacimiento")
	l_domicilio = l_rs("domicilio")
	l_telefono = l_rs("telefono")
end if
l_rs.close


 %>

<body leftmargin="0" rightmargin="0" topmargin="0" bottommargin="0" onload="//javascript:parent.Buscar();">


 <a align="center" class=sidebtnSHW href="javascript:imprSelec('seleccion')"><img  src="../shared/images/print-icon_72.png" border="0" title="Imprimir"></a>
  <div id="seleccion">
<table>

    <tr>
        <td align="center" colspan="6"><img  src="/turnos/images/megavision.jpg" border="0"></a> </td>
    </tr>
    <tr>
        <td align="center" colspan="6">&nbsp;</td>
    </tr>		
    <tr>
        <td align="center" colspan="6">___________________________________________________________________________________________</td>
    </tr>		
	<tr>
        <td  colspan="6" align="left" ><h3>Datos del Paciente:&nbsp;<%= l_paciente %>&nbsp;&nbsp;&nbsp;&nbsp;</h3></td>
	 </tr>	
	<tr>
        <td  colspan="6" align="left" ><h3>Documento N:&nbsp;<%= l_dni %>&nbsp;&nbsp;&nbsp;&nbsp;</h3></td>
	 </tr>	
	<tr>
        <td  colspan="6" align="left" ><h3>Domicilio:&nbsp;<%= l_domicilio %>&nbsp;&nbsp;&nbsp;&nbsp;</h3></td>
	 </tr>	
	<tr>
        <td  colspan="6" align="left" ><h3>Telefono:&nbsp;<%= l_telefono %>&nbsp;&nbsp;&nbsp;&nbsp;</h3></td>
	 </tr>	
	<tr>
        <td  colspan="6" align="left" >&nbsp;</td>
	 </tr>	 
	<tr>
        <td  colspan="6" align="left" >&nbsp;</td>
	 </tr>		
	 	 	
					 	 

	
    <tr>
        <td colspan="6"><h4>Observaciones</h4></td>
    </tr>	   
    <tr>
        <td colspan="6"><h3><%= l_detalle %></h3></td>
    </tr>	
	<tr>
        <td  colspan="6" align="left" >&nbsp;</td>
	 </tr>	
	<tr>
        <td  colspan="6" align="left" >&nbsp;</td>
	 </tr>	
	<tr>
        <td  colspan="6" align="left" >&nbsp;</td>
	 </tr>		 	 	
	
	<tr>
        <td  colspan="6" align="center" ><h3>Medico:&nbsp;<%= l_medico %>&nbsp;&nbsp;&nbsp;&nbsp;</h3></td>
	</tr>		
	
</div>
    <%

set l_rs = Nothing
cn.Close
set cn = Nothing
%>

</table>
<form name="datos" method="post">
<input type="hidden" name="cabnro" value="0">
<input type="hidden" name="idturno" value="0">
<input type="hidden" name="orden" value="<%= l_orden %>">
<input type="hidden" name="filtro" value="<%= l_filtro %>">
</form>
</body>
</html>

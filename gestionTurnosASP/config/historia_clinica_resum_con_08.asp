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
dim l_fecha
dim l_Medico_imagen

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
<meta http-equiv="Content-Type" content="text/html; charset=iso-8859-1" />
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

function Nombre_Mes (Mes)

	select case Mes
		case "1":
			Nombre_Mes = "Enero"
		case "2":
			Nombre_Mes = "Febrero"
		case "3":
			Nombre_Mes = "Marzo"
		case "4":
			Nombre_Mes = "Abril"
		case "5":
			Nombre_Mes = "Mayo"
		case "6":
			Nombre_Mes = "Junio"
		case "7":
			Nombre_Mes = "Julio"
		case "8":
			Nombre_Mes = "Agosto"
		case "9":
			Nombre_Mes = "Septiembre"
		case "10":
			Nombre_Mes = "Octubre"
		case "11":
			Nombre_Mes = "Noviembre"
		case "12":
			Nombre_Mes = "Diciembre"
		
		end select	
									

end function

function Formatear_Texto (Cadena)

Dim l_rs

Set l_rs = Server.CreateObject("ADODB.RecordSet")

l_sql = "SELECT * "
l_sql = l_sql & " FROM empresa "
l_sql = l_sql & " where empresa.id = " & Session("empnro")   
rsOpen l_rs, cn, l_sql, 0 
if l_rs.eof then
	response.write Cadena
Else  
	if l_rs("hist_clin_bold") = "Y" then
		response.write "<FONT SIZE='" & l_rs("hist_clin_size") & "' FACE='" & l_rs("hist_clin_face") & "'><b>" & Cadena & "</b></FONT>"
	else
		response.write "<FONT SIZE='" & l_rs("hist_clin_size") & "' FACE='" & l_rs("hist_clin_face") & "'>" & Cadena & "</FONT>"
	end if
End If 

l_rs.close

end function

Set l_rs = Server.CreateObject("ADODB.RecordSet")

' Obtengo la cantidad de turnos simultaneos del Recurso Reservable
l_sql = "SELECT  historia_clinica_resumida.* , clientespacientes.apellido, clientespacientes.nombre, recursosreservables.descripcion , clientespacientes.dni, clientespacientes.fechanacimiento , clientespacientes.domicilio, clientespacientes.telefono , recursosreservables.firma  Medico_imagen "
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
	l_fecha = l_rs("fecha")
	l_Medico_imagen = l_rs("Medico_imagen")
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
       <td align="right" colspan="6"><%= Formatear_Texto ( "Bah&iacute;a Blanca, " &  day(l_fecha) & " " & Nombre_Mes( month(l_fecha)) & " de " & year(l_fecha) )%></td>
   </tr>		
   <tr>
       <td align="center" colspan="6">&nbsp;</td>
   </tr>	   
    <tr>
        <td align="center" colspan="6">___________________________________________________________________________________________</td>
    </tr>		
	<tr>
        <td  colspan="6" align="left" > <%= Formatear_Texto ( "Datos del Paciente:&nbsp;" & l_paciente ) %></td>
	 </tr>	
	<tr>
        <td  colspan="6" align="left" > <%= Formatear_Texto ( "Documento N:&nbsp;" & l_dni ) %></td>
	 </tr>	
	<tr>
        <td  colspan="6" align="left" > <%= Formatear_Texto ( "Domicilio:&nbsp;" & l_domicilio ) %></td>
	 </tr>	
	<tr>
        <td  colspan="6" align="left" > <%= Formatear_Texto ( "Telefono:&nbsp;" & l_telefono )%></td>
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
        <td colspan="6"> <%= Formatear_Texto ( l_detalle ) %></td>
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
	
	<% if not isnull(l_Medico_imagen) then %> 
    <tr>
        <td align="center" colspan="6"><img  src="/turnos/images/<%= l_Medico_imagen %>" border="0"> </a> </td>
    </tr>	
	<% End If %> 	 	 	
	
	<tr>
        <td  colspan="6" align="center" > <%= Formatear_Texto ( "Medico:&nbsp;" & l_medico )%></td>
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

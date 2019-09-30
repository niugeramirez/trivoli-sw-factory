
<% Option Explicit %>
<!--#include virtual="/turnos/shared/db/conn_db.inc"-->
<!--#include virtual="/turnos/shared/inc/fecha.inc"-->
<% 

on error goto 0

'ADO
Dim l_sql
Dim l_rs
Dim l_orden
Dim l_fechadesde
Dim l_fechahasta
Dim l_idos
Dim l_tipo
Dim l_params

l_orden = " ORDER BY  visitas.fecha, osnombre, nombremedico "

l_fechadesde = request("qfechadesde")
'Response.write l_fechadesde
l_fechahasta = request("qfechahasta")
l_idos = request("idos")
l_tipo = request("tipo")
l_params = l_fechadesde & "," & l_fechahasta & "," & l_idos & "," & l_tipo
                        
Set l_rs = Server.CreateObject("ADODB.RecordSet")

l_sql = "SELECT  visitas.fecha, practicas.descripcion nombrepractica , recursosreservables.descripcion nombremedico, isnull(practicasrealizadas.id,0) practicasrealizadasid , practicasrealizadas.precio , visitas.flag_ausencia  " 
l_sql = l_sql & " ,  clientespacientes.apellido, clientespacientes.nombre"
l_sql = l_sql & " ,  obrassociales.descripcion osnombre"
l_sql = l_sql & " ,  ( select min(ospago.descripcion) from pagos LEFT JOIN obrassociales ospago ON ospago.id = pagos.idobrasocial where pagos.idpracticarealizada = practicasrealizadas.id ) ospago "
l_sql = l_sql & " FROM visitas "
l_sql = l_sql & " INNER JOIN practicasrealizadas ON practicasrealizadas.idvisita = visitas.id "
l_sql = l_sql & " LEFT JOIN recursosreservables ON recursosreservables.id = visitas.idrecursoreservable "
l_sql = l_sql & " LEFT JOIN clientespacientes ON clientespacientes.id = visitas.idpaciente "
l_sql = l_sql & " LEFT JOIN obrassociales ON obrassociales.id = clientespacientes.idobrasocial "
l_sql = l_sql & " LEFT JOIN practicas ON practicas.id = practicasrealizadas.idpractica "
l_sql = l_sql & " WHERE  visitas.fecha  >= " & cambiafecha(l_fechadesde,"YMD",true) 
l_sql = l_sql & " AND  visitas.fecha <= " & cambiafecha(l_fechahasta,"YMD",true) 
l_sql = l_sql & " AND  isnull(visitas.flag_ausencia,0) <> -1" 
l_sql = l_sql & " and visitas.empnro = " & Session("empnro")   

if l_tipo <> "T" then
	if l_tipo = "O" then
		l_sql = l_sql & " AND clientespacientes.afiliado_obligatorio = 'S'"
	else
		l_sql = l_sql & " AND (clientespacientes.afiliado_obligatorio IS NULL OR clientespacientes.afiliado_obligatorio <> 'S')"
	end if
end if

if l_idos <> "0" then
	l_sql = l_sql &" AND exists ( select (ospago.id) from pagos LEFT JOIN obrassociales ospago ON ospago.id = pagos.idobrasocial where pagos.idpracticarealizada = practicasrealizadas.id and ospago.id IN "& l_idos & ")" 
end if
l_sql = l_sql & " " & l_orden

'rsOpen l_rs, cn, l_sql, 0 

'do until l_rs.eof
'	l_rs.MoveNext
'loop

'l_rs.Close
set l_rs = Nothing
cn.Close
set cn = Nothing
%>
<html>
<head>
<link href="/turnos/ess/shared/css/tables_gray.css" rel="StyleSheet" type="text/css">
<meta http-equiv="Content-Type" content="text/html; charset=iso-8859-1">
<title>Liquidacion Visitas</title>
</head>
<script src="/turnos/shared/js/fn_valida.js"></script>
<script src="/turnos/shared/js/fn_fechas.js"></script>
<script src="/turnos/shared/js/fn_ayuda.js"></script>
<script src="/turnos/shared/js/fn_windows.js"></script>
<script src="/turnos/shared/js/fn_numeros.js"></script>

<script>
function Validar_Formulario(){

if (document.datos.factura.value == ""){
	alert("Debe ingresar el Numero de Factura.");
	document.datos.factura.focus();
	return;
}
else{
	var d=document.datos;
	document.valida.location = "rep_visitas_por_OS_rep_06.asp";
	}	
}

function valido(){
	document.datos.submit();
}

function invalido(texto){
	alert(texto);
	document.datos.coudes.focus();
}

function Nuevo_Dialogo(w_in, pagina, ancho, alto)
{
 return w_in.showModalDialog(pagina,'', 'center:yes;dialogWidth:' + ancho.toString() + ';dialogHeight:' + alto.toString() + ';');
}
function Ayuda_Fecha(txt)
{
 var jsFecha = Nuevo_Dialogo(window, '/turnos/shared/js/calendar.html', 16, 15);

 if (jsFecha == null) txt.value = ''
 else txt.value = jsFecha;
}

function Mayuscula(cadena){

	cadena.value = cadena.value.toUpperCase();
}

</script>
<% 

%>
<body leftmargin="0" rightmargin="0" topmargin="0" bottommargin="0" onload="javascript:document.datos.factura.focus();">
<form name="datos" action="rep_visitas_por_OS_rep_03.asp" method="post" target="valida">
<input type="Hidden" name="fechadesde" value="<%= l_fechadesde %>">
<input type="Hidden" name="fechahasta" value="<%= l_fechahasta %>">
<input type="Hidden" name="tipo" value="<%= l_tipo %>">
<input type="Hidden" name="idos" value="<%= l_idos %>">

<table cellspacing="0" cellpadding="0" border="0" width="100%" height="100%">
<tr>
    <td class="th2" nowrap>&nbsp;</td>
	<td class="th2" align="right">
		<a class=sidebtnHLP href="Javascript:ayuda('<%= Request.ServerVariables("SCRIPT_NAME")%>');">Ayuda</a>
	</td>
</tr>
<tr>
	<td colspan="2" height="100%">
		<table border="0" cellspacing="0" cellpadding="0">
			<tr>
				<td>
					<table cellspacing="0" cellpadding="0" border="0">
					<tr>
					    <td align="center"><b>Numero de Factura (*):</b></td>
						<td>
							<input type="text" name="factura" size="8" maxlength="8">							
						</td>
					    <td align="center"><b>Registrar Envio:</b></td>
						<td>
							<input  type=checkbox name="registrar_envio"> 
						</td>
					</tr>	
					</table>
				</td>
			</tr>
		</table>
	</td>
</tr>
<tr>
    <td colspan="2" align="right" class="th">
		<a class=sidebtnABM href="Javascript:Validar_Formulario()">Aceptar</a>
		<a class=sidebtnABM href="Javascript:window.close()">Cancelar</a>
	</td>
</tr>

</table>
<iframe name="valida" style="visibility=hidden;" src="" width="100%" height="100%"></iframe> 
</form>
</body>
</html>

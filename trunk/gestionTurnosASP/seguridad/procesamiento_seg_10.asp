<% Option Explicit %>
<!--#include virtual="/serviciolocal/shared/db/conn_db.inc"-->

<% 
'Archivo: procesamiento_seg_10.asp
'Descripción:
'Autor : Lisandro Moro
'Fecha : 10/03/2005
'Modificado:

Dim l_bpronro
Dim l_bprchora
Dim l_bprcfecha
Dim l_btprcdesabr
Dim l_bprcestado
Dim l_bprcprogreso
Dim l_iduser
Dim l_bprcfecdesde
Dim l_bprcfechasta
Dim l_bprcUrgente
Dim l_bprcfecInicioEj
Dim l_bprcfecFinEj
Dim l_bprcHoraInicioEj
Dim l_bprcHoraFinEj
Dim l_bprcterminar
Dim l_cantEmp
Dim l_rs
Dim l_sql
Dim l_tipo
Dim l_bprcparam
Dim l_btprcnro

l_bpronro = request("bpronro")

%>
<html>
<head>
<link href="/serviciolocal/shared/css/tables3.css" rel="StyleSheet" type="text/css">
<meta http-equiv="Content-Type" content="text/html; charset=iso-8859-1">
<title><%= Session("Titulo")%>Procesamiento - Gesti&oacute;n de Tiempos - RHPro &reg;</title>
</head>
<script src="/serviciolocal/shared/js/fn_ayuda.js"></script>
<script src="/serviciolocal/shared/js/fn_fechas.js"></script>
<script>
function Validar_Formulario(){
	if (document.datos.bprcterminar.checked){
		if (confirm('¿ Desea eliminar el registro seleccionado ?') == true)
			document.datos.submit();
		else
			document.datos.bprcterminar.checked = false;
	}	
	else
		if ((document.datos.bprcfecdesde.value == "") && (document.datos.bprcfechasta.value == ""))
			document.datos.submit();	
		else
			if (validarfecha(document.datos.bprcfecdesde) && validarfecha(document.datos.bprcfechasta)){
				document.datos.submit();
			}
}

function Nuevo_Dialogo(w_in, pagina, ancho, alto){
	return w_in.showModalDialog(pagina,'', 'center:yes;dialogWidth:' + ancho.toString() + ';dialogHeight:' + alto.toString() + ';');
}

function Ayuda_Fecha(txt){
	var jsFecha = Nuevo_Dialogo(window, '/serviciolocal/shared/js/calendar.html', 16, 15);
	if (jsFecha == null)
		txt.value = ''
	else
		txt.value = jsFecha;
}
</script>
<% 
Set l_rs = Server.CreateObject("ADODB.RecordSet")
l_sql = "SELECT bpronro, bprchora, bprcfecha,btprcdesabr,bprcestado,bprcprogreso,iduser, bprcfecdesde "
l_sql = l_sql & ", bprcfechasta, bprcUrgente, bprcfecInicioEj, bprcfecFinEj, bprcHoraInicioEj, bprcHoraFinEj, bprcterminar, bprcparam, batch_proceso.btprcnro "
l_sql = l_sql & "FROM batch_proceso INNER JOIN batch_tipproc ON batch_tipproc.btprcnro = batch_proceso.btprcnro "
l_sql = l_sql & "where batch_proceso.bpronro="& l_bpronro

rsOpen l_rs, cn, l_sql, 0 
if not l_rs.eof then
	l_bprchora = l_rs("bprchora")
	l_bprcfecha = l_rs("bprcfecha")
	l_btprcdesabr = l_rs("btprcdesabr")
	l_bprcestado = l_rs("bprcestado")
	l_bprcprogreso = l_rs("bprcprogreso")
	l_iduser = l_rs("iduser")
	l_bprcfecdesde = l_rs("bprcfecdesde")
	l_bprcfechasta = l_rs("bprcfechasta")
	l_bprcUrgente = l_rs("bprcUrgente")
	l_bprcfecInicioEj = l_rs("bprcfecInicioEj")
	l_bprcfecFinEj = l_rs("bprcfecFinEj")
	l_bprcHoraInicioEj = l_rs("bprcHoraInicioEj")
	l_bprcHoraFinEj = l_rs("bprcHoraFinEj")
	l_bprcterminar = l_rs("bprcterminar")
	l_bprcparam = l_rs("bprcparam")
	l_btprcnro = l_rs("btprcnro")
end if
l_rs.Close

'l_sql = "SELECT COUNT(ternro) as cant from batch_empleado where bpronro="& l_bpronro
'rsOpen l_rs, cn, l_sql, 0 
'if not l_rs.eof then
'	l_cantEmp = l_rs("cant")
l_cantEmp = 0
'end if
'l_rs.Close

set l_rs = nothing

%>
<body leftmargin="0" rightmargin="0" topmargin="0" bottommargin="0">
<form name="datos" action="procesamiento_seg_12.asp" method="post">
<input type="Hidden" name="bprcparam" value="<%= l_bprcparam%>">
<input type="Hidden" name="btprcnro" value="<%= l_btprcnro%>">

<table cellspacing="0" cellpadding="0" border="0" width="100%" height="100%">
	<tr>
		<td class="th2">Datos del Proceso</td>
		<td class="th2" align="right">
			<a class=sidebtnHLP href="Javascript:ayuda('<%= Request.ServerVariables("SCRIPT_NAME")%>');">Ayuda</a>
		</td>
	</tr>
	<tr>
		<td colspan="2" width="100%" height="100%">
			<table cellspacing="0" cellpadding="0" border="0" width="100%" height="100%">
				<tr>
					<td width="50%"></td>
					<td>
						<table cellpadding="0" cellspacing="0" border="0">
							<tr>
							    <td align="right"><b>Número:</b></td>
								<td align="left"><input type="text" name="bpronro" size="5" readonly class="deshabinp" value='<%= l_bpronro %>'></td>
							    <td align="right"><b>Tipo:</b></td>
								<td align="left"><input type="text" name="btprcdesabr" size="25" readonly class="deshabinp" value='<%= l_btprcdesabr %>'></td>
							</tr>
							<tr>
							    <td align="right"><b>Fecha:</b></td>
								<td align="left"><input type="text" name="bprcfecha" size="10" readonly class="deshabinp" value='<%= l_bprcfecha %>'></td>
							    <td align="right"><b>Hora:</b></td>
								<td align="left">
									<input type="text" name="bprchora" size="8" readonly class="deshabinp" value='<%=l_bprchora%>'>
								</td>
							</tr>
							<tr>
							    <td align="right"><b>Estado:</b></td>
								<td align="left" colspan="2"><input type="text" name="bprcestado" size="25" readonly class="deshabinp" value='<%= l_bprcestado %>'></td>
								<td>
								<%if (l_bprcestado = "Incompleto") OR (InStr(l_bprcestado,"Abortado") > 0) then%>
									<input type="checkbox" name="estadopasar"><b>Pasar a Pendiente</b>
								<%end if%>	
								</td>
							</tr>
							<tr>
							    <td align="right"><b>Progreso:</b></td>
								<td align="left"><input type="text" name="bprcprogreso" size="3" readonly class="deshabinp" value='<%= l_bprcprogreso %>'></td>
							    <td align="right" nowrap><b>Cant. Emp.:</b></td>
								<td align="left"><input type="text" name="cantEmp" size="5" readonly class="deshabinp" value='<%= l_cantEmp %>'></td>
							</tr>
							<tr>
							    <td align="right"><b>Desde:</b></td>
								<td nowrap>
								<input  type="text" name="bprcfecdesde" size="10" maxlength="10" value="<%= l_bprcfecdesde %>" <%if l_bprcestado <> "Pendiente" then%>readonly class="deshabinp"<%end if%>>
								<%if l_bprcestado = "Pendiente" then%>
									<a href="Javascript:Ayuda_Fecha(document.datos.bprcfecdesde);"><img src="/serviciolocal/shared/images/cal.gif" border="0"></a>	
								<%end if%>	
								</td>
							    <td align="right"><b>Hasta:</b></td>
								<td nowrap>
								<input type="text" <%if l_bprcestado <> "Pendiente" then%>readonly class="deshabinp"<%end if%> name="bprcfechasta" size="10" maxlength="10" value="<%= l_bprcfechasta %>" >
								<%if l_bprcestado = "Pendiente" then%>
								<a href="Javascript:Ayuda_Fecha(document.datos.bprcfechasta);"><img src="/serviciolocal/shared/images/cal.gif" border="0"></a>	
								<%end if%>
								</td>
							</tr>
							<tr>
							    <td colspan="4" height="10"></td>
							</tr>
							<tr>
							    <td align="right" valign="middle"><b>Prioridad:</b></td>
							    <td colspan="3">
									<table cellpadding="0" cellpadding="0" border="0">
									    <td align="right"></td>
										<td><input type="checkbox" <%if l_bprcestado <> "Pendiente" then%>disabled <%end if%> 
											<%if l_bprcUrgente then%>checked<%end if%> name="bprcUrgente"><b>Urgente</b></td>
									    <td align="right"></td>
										<td><input type="checkbox" <%if (l_bprcestado <> "Pendiente" and l_bprcestado <> "Procesando" and l_bprcestado <> "No Responde") then%>disabled <%end if%> 
											<%if l_bprcterminar then%>checked<%end if%> name="bprcterminar"><b>Terminar</b></td>
									</table>
								</td>
							</tr>
							<tr>
							    <td colspan="4" height="10" align="left"></td>
							</tr>
							<tr>
								<td align="right" nowrap><b>Fecha Inicio:</b></td>
								<td align="left">
									<input type="text" name="bprcfecInicioEj" size="10" readonly class="deshabinp" value='<%= l_bprcfecInicioEj %>'>
								</td>
							    <td align="right" nowrap><b>Hora Inicio:</b></td>
								<td align="left">
									<input type="text" name="bprcHoraInicioEj" size="8" readonly class="deshabinp" value='<%= l_bprcHoraInicioEj %>'>
								</td>
							</tr>
							<tr>
								<td align="right" nowrap><b>Fecha Fin:</b></td>
								<td align="left">
									<input type="text" name="bprcfecFinEj" size="10" readonly class="deshabinp" value='<%= l_bprcfecFinEj %>'>
								</td>
							    <td align="right"><b>Hora Fin:</b></td>
								<td align="left">
									<input type="text" name="bprcHoraFinEj" size="8" readonly class="deshabinp" value='<%= l_bprcHoraFinEj %>'>
								</td>
							</tr>
						</table>
					</td>
					<td width="50%"></td>
				</tr>
			</table>
		</td>
	</tr>
	<tr>
		<td align="right" class="th2" colspan="2">
			<a class=sidebtnABM href="Javascript:Validar_Formulario()">Aceptar</a>
			<a class=sidebtnABM href="Javascript:window.close()">Cancelar</a>
		</td>
	</tr>
</table>
</form>
</body>
</html>

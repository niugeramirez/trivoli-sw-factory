
<% Option Explicit %>
<!--#include virtual="/turnos/shared/db/conn_db.inc"-->
<% 

on error goto 0

'Datos del formulario

dim l_id
dim l_apellido
dim l_nombre  
dim l_dni     
dim l_domicilio
dim l_idobrasocial
'ADO
Dim l_tipo
Dim l_sql
Dim l_rs

dim l_idpractica 
dim l_idsolicitadapor
dim l_precio

Dim l_idvisita
Dim l_idpracticarealizada

l_tipo = request.querystring("tipo")
l_idvisita = request("cabnro")

l_idpracticarealizada = request("idpracticarealizada")

%>
<html>
<head>
<link href="/turnos/ess/shared/css/tables_gray.css" rel="StyleSheet" type="text/css">
<!--<link href="/turnos/shared/css/tables_gray.css" rel="StyleSheet" type="text/css">-->
<meta http-equiv="Content-Type" content="text/html; charset=iso-8859-1">
<title>Agregar Practica</title>
</head>
<script src="/turnos/shared/js/fn_valida.js"></script>
<script src="/turnos/shared/js/fn_fechas.js"></script>
<script src="/turnos/shared/js/fn_ayuda.js"></script>
<script src="/turnos/shared/js/fn_windows.js"></script>
<script src="/turnos/shared/js/fn_numeros.js"></script>
<script>
function Validar_Formulario(){

if (document.datos.practicaid.value == "0"){
	alert("Debe ingresar la Practica.");
	document.datos.practicaid.focus();
	return;
}

valido();
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


function ppa(){
	alert();
}

</script>
<% 
select Case l_tipo
	Case "A":
			l_idpractica = 0
			l_idsolicitadapor = 0
			l_precio = 0
	Case "M":
		Set l_rs = Server.CreateObject("ADODB.RecordSet")
		l_id = request.querystring("cabnro")
		l_sql = "SELECT * "
		l_sql = l_sql & " FROM practicasrealizadas "
		l_sql  = l_sql  & " WHERE id = " & l_idpracticarealizada
		rsOpen l_rs, cn, l_sql, 0 
		if not l_rs.eof then
			l_idpractica = l_rs("idpractica")
			l_idsolicitadapor = l_rs("idsolicitadapor") 
			l_precio = l_rs("precio")
		end if
		l_rs.Close
end select
%>
<body leftmargin="0" rightmargin="0" topmargin="0" bottommargin="0" onload="javascript:document.datos.motivo.focus();">
<form name="datos" action="AgregarPractica_con_03.asp?tipo=<%= l_tipo %>" method="post" target="valida">
<input type="hidden" name="idvisita" value="<%= l_idvisita %>">
<input type="hidden" name="idpracticarealizada" value="<%= l_idpracticarealizada %>">



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
					<!-- 
					<tr>
						<td align="left" colspan="4" style="font-size:20"  >
							Servicio Local: <b><%'= l_serdes %><b>				
						</td>																	
					</tr>  -->						
											
					<tr>
						<td  align="right" nowrap><b>Practica (*): </b></td>
						<td colspan="3"><select name="practicaid" size="1" style="width:200;">
								<option value=0 selected>Seleccione una Practica</option>
								<%Set l_rs = Server.CreateObject("ADODB.RecordSet")
								l_sql = "SELECT  * "
								l_sql  = l_sql  & " FROM practicas "
								l_sql  = l_sql  & " ORDER BY descripcion "
								rsOpen l_rs, cn, l_sql, 0
								do until l_rs.eof		%>	
								<option value= <%= l_rs("id") %> > 
								<%= l_rs("descripcion") %> </option>
								<%	l_rs.Movenext
								loop
								l_rs.Close %>
							</select>
							<script>document.datos.practicaid.value="<%= l_idpractica %>"</script>
						</td>					
					</tr>	
					
					<tr>
						<td  align="right" nowrap><b>Solicitado por : </b></td>
						<td colspan="3"><select name="idrecursoreservable" size="1" style="width:200;">
								<option value=0 selected>Ningun Profesional</option>
								<%Set l_rs = Server.CreateObject("ADODB.RecordSet")
								l_sql = "SELECT  * "
								l_sql  = l_sql  & " FROM recursosreservables "
								l_sql  = l_sql  & " ORDER BY descripcion "
								rsOpen l_rs, cn, l_sql, 0
								do until l_rs.eof		%>	
								<option value= <%= l_rs("id") %> > 
								<%= l_rs("descripcion") %> </option>
								<%	l_rs.Movenext
								loop
								l_rs.Close %>
							</select>
							<script>document.datos.idrecursoreservable.value="<%= l_idsolicitadapor %>"</script>							
						</td>					
					</tr>		
					<% if l_tipo = "M" then %>
					<tr>
					    <td align="right"><b>Precio:</b></td>
						<td colspan="3">
							<input type="text" name="precio" size="20" maxlength="20" value="<%= l_precio %>">
						</td>
					</tr>		
					<% End If %>							
									
					
					<!--
					<tr>
					    <td align="right" ><b>Fec. Nac.:</b></td>
						<td align="left"  >
						    <input type="text" name="legfecnac" size="10" maxlength="10" value="<%'= l_legfecnac %>">
							<a href="Javascript:Ayuda_Fecha(document.datos.legfecnac)"><img src="/turnos/shared/images/cal.gif" border="0"></a>
						</td>
						<td align="right"><b>Teléfono:</b></td>
						<td>
							<input type="text" name="legtel" size="20" maxlength="20" value="<%'= l_legtel %>">
						</td>						
					</tr>
					-->
					<!-- 
					<tr>
						<td  align="right" nowrap><b>Derecho Vulnerado: </b></td>
						<td colspan="3"><select name="pronro" size="1" style="width:150;">
								<option value=0 selected>Todos</option>
								<%'Set l_rs = Server.CreateObject("ADODB.RecordSet")
								'l_sql = "SELECT  * "
								'l_sql  = l_sql  & " FROM ser_problematica "
								'l_sql  = l_sql  & " ORDER BY prodes "
								'rsOpen l_rs, cn, l_sql, 0
								'do until l_rs.eof		%>	
								<option value= <%'= l_rs("pronro") %> > 
								<%'= l_rs("prodes") %> (<%'=l_rs("pronro")%>) </option>
								<%'	l_rs.Movenext
								'loop
								'l_rs.Close %>
							</select>
							<script>document.datos.pronro.value= "<%'= l_pronro %>"</script>
						</td>					
					</tr>
					<tr>
					    <td align="right"><b>Madre - Apellido y Nombre:</b></td>
						<td>
							<input type="text" name="legapenommad" size="20" maxlength="20" value="<%'= l_legapenommad %>">
						</td>
						<td align="right"><b>Dom:</b></td>						
						<td>
							<input type="text" name="legdommad" size="20" maxlength="20" value="<%'= l_legdommad %>">
							<b>Tel:</b> <input type="text" name="legtelmad" size="10" maxlength="20" value="<%'= l_legtelmad %>">						
						</td>							
					</tr>																				
					<tr>
					    <td align="right"><b>Padre - Apellido y Nombre:</b></td>
						<td>
							<input type="text" name="legapenompad" size="20" maxlength="20" value="<%'= l_legapenompad  %>">
						</td>
						<td align="right"><b>Dom:</b></td>												
						<td>
							<input type="text" name="legdompad" size="20" maxlength="20" value="<%'= l_legdompad %>">
							<b>Tel:</b> <input type="text" name="legtelpad" size="10" maxlength="20" value="<%'= l_legtelpad %>">
						</td>						
					</tr>					
					<tr>
					    <td align="right"><b>Instituciones Intervinientes:</b></td>
						<td colspan="3">
							<input type="text" name="legins" size="80" maxlength="20" value="<%'= l_legins %>">
						</td>
					</tr>																				
					<tr>
					    <td align="right"><b>Instituciones Educativas:</b></td>
						<td colspan="3">
							<input type="text" name="leginsedu" size="80" maxlength="20" value="<%'= l_leginsedu %>">
						</td>
					</tr>																									
					<tr>
					    <td align="right"><b>Cobertura Social de la Familia:</b></td>
						<td colspan="3">
							<input type="text" name="legcobsoc" size="80" maxlength="20" value="<%'= l_legcobsoc %>">
						</td>
					</tr>																														
					<tr>
					    <td align="right"><b>Estrategias de Intervención:</b></td>
						<td colspan="3">
							<input type="text" name="legabo" size="80" maxlength="20" value="<%'= l_legabo %>">
						</td>
					</tr>					
					<tr>
						<td align="right"><b>Medidas Protección:</b></td>
						<td colspan="3"><select name="mednro" size="1" style="width:150;">
								<option value=0 selected>&nbsp;</option>
								<%'Set l_rs = Server.CreateObject("ADODB.RecordSet")
								'l_sql = "SELECT  * "
								'l_sql  = l_sql  & " FROM ser_medida "
								'l_sql  = l_sql  & " ORDER BY meddes "
								'rsOpen l_rs, cn, l_sql, 0
								'do until l_rs.eof		%>	
								<option value= <%'= l_rs("mednro") %> > 
								<%'= l_rs("meddes") %> (<%'=l_rs("mednro")%>) </option>
								<%'	l_rs.Movenext
								'loop
								'l_rs.Close %>
							</select>
							<script>document.datos.mednro.value= "<%'= l_mednro %>"</script>
						</td>					
					</tr>					
					 -->						
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
<%
set l_rs = nothing
cn.Close
set cn = nothing
%>
</body>
</html>

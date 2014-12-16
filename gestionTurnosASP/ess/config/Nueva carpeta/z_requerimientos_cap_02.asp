<% Option Explicit %>
<!--#include virtual="/ticket/shared/db/conn_db.inc"-->
<!--
Archivo: requerimientos_cap_02.asp
Descripción: Abm de requerimientos
Autor : Lisandro Moro
Fecha: 11/11/2003
Modificado: Raul Chinestra - Se agrego la condicion que la fecha pretendida sea mayor que la fecha del requerimiento.
MOdificado: lisandro Moro - 17/05/2004 - Se Agrego Firma.
MOdificado: lisandro Moro - 22/11/2004 - Correccion en la ventana pq se veia cortada.
-->
<% 
'on error goto 0
'Datos del formulario
'pednro, peddesabr, peddesext, estpednro, modnro, pedfecped, pedfecpret, peddurpredias, peddurprethora, pedpers, pedsolpor, pedrelpor, pedprioridad, pedmotprio
Dim l_pednro
Dim l_peddesabr
Dim l_peddesext
Dim l_estpednro
Dim l_modnro
Dim l_pedfecped
Dim l_pedfecpret
Dim l_peddurpredias
Dim l_peddurprethora
Dim l_pedpers
Dim l_pedsolpor
Dim l_pedrelpor
Dim l_pedprioridad
Dim l_pedmotprio

'ADO
Dim l_tipo
Dim l_sql
Dim l_sql1
Dim l_sql2
Dim l_rs
Dim l_rs1
Dim l_rs2
Dim l_temnro

l_tipo = request.querystring("tipo")
%>
<html>
<head>
<link href="/ticket/shared/css/tables3.css" rel="StyleSheet" type="text/css">
<meta http-equiv="Content-Type" content="text/html; charset=iso-8859-1">
<title><%= Session("Titulo")%>Requerimientos - Capacitación - RHPro &reg;</title>
</head>
<script src="/ticket/shared/js/fn_ayuda.js"></script>
<script src="/ticket/shared/js/fn_windows.js"></script>
<script src="/ticket/shared/js/fn_fechas.js"></script>
<script src="/ticket/shared/js/fn_numeros.js"></script>
<% 

'-----------------------------------------------------------------------------------------------------
Set l_rs = Server.CreateObject("ADODB.RecordSet")
select Case l_tipo
	Case "A":
		'l_pednro = ""
		l_peddesabr = ""
		l_peddesext = ""
		l_estpednro = ""
		l_modnro = ""
		l_pedfecped = date
		l_pedfecpret = date
		l_peddurpredias = ""
		l_peddurprethora = ""
		l_pedpers = ""
		l_pedsolpor = ""
		l_pedrelpor = ""
		l_pedprioridad = ""
		l_pedmotprio = ""
		
	Case "M":
		l_temnro = request.querystring("cabnro")
		l_sql = "SELECT pednro, peddesabr, peddesext, estpednro, modnro, pedfecped, pedfecpret, peddurpredias, peddurprethora, pedpers, pedsolpor, pedrelpor, pedprioridad, pedmotprio"
		l_sql = l_sql & " FROM cap_pedido"
		l_sql  = l_sql  & " WHERE pednro = " & l_temnro
		rsOpen l_rs, cn, l_sql, 0 
		if not l_rs.eof then
			
		l_pednro = l_rs("pednro")
		l_peddesabr = l_rs("peddesabr")
		l_peddesext = l_rs("peddesext")
		l_estpednro = l_rs("estpednro")
		l_modnro = l_rs("modnro")
		l_pedfecped = l_rs("pedfecped")
		l_pedfecpret = l_rs("pedfecpret")
		l_peddurpredias = l_rs("peddurpredias")
		l_peddurprethora = l_rs("peddurprethora")
		l_pedpers = l_rs("pedpers")
		l_pedsolpor = l_rs("pedsolpor")
		l_pedrelpor = l_rs("pedrelpor")
		l_pedprioridad = l_rs("pedprioridad")
		l_pedmotprio = l_rs("pedmotprio")

		end if
		l_rs.Close
end select

'-------------Firmas-----------
 Dim l_tipAutorizacion  'Es el tipo del circuito de firmas
 Dim l_HayAutorizacion  'Es para ver si las autorizaciones estan activas
 Dim l_PuedeVer         'Es para ver si las autorizaciones estan activas
 
 l_sql = "select cystipo.* from cystipo "
 l_sql = l_sql & "where (cystipo.cystipact = -1) and cystipo.cystipnro = 8 "
 
 rsOpen l_rs, cn, l_sql, 0 
 
 l_HayAutorizacion = not l_rs.eof
 
 if not l_rs.eof then
 	l_tipAutorizacion = 8
 end if 
 l_rs.close
 
 if l_HayAutorizacion AND (l_tipo = "M") then
 
	l_sql = "select cysfirautoriza, cysfirsecuencia, cysfirdestino from cysfirmas "
	l_sql = l_sql & "where cysfirmas.cystipnro = " & l_tipAutorizacion & " and cysfirmas.cysfircodext = '" & l_temnro & "' " 
	l_sql = l_sql & "order by cysfirsecuencia desc"
 
	rsOpen l_rs, cn, l_sql, 0 
 
	l_PuedeVer = False
 
	if not l_rs.eof then
 		if (l_rs("cysfirautoriza") = session("UserName")) or (l_rs("cysfirdestino") = session("UserName")) then 
	   		'Es una modificación del ultimo o es el nuevo que autoriza 
    		l_PuedeVer = True 
    	end if
 	end if
	l_rs.close
 	If not l_PuedeVer then
    	response.write "<script>alert('No esta autorizado a ver o modificar este registro.');window.close()</script>"
		response.end
 	End if
 End if
'-------------Firmas-----------
%>
<script>
function Validar_Formulario(){
//var selfil;
//var nselfil;
if (document.datos.peddesabr.value == "") {
	alert("Debe ingresar la Descripción Abreviada.");document.datos.peddesabr.focus();
}else{
	<% if l_tipo = "A" then %>
	if (document.datos.curnro.value == 0) {
    	alert("Debe Seleccionar un Curso");document.datos.curnro.focus();
	}else{
		if (document.datos.micadena.value == '' || document.datos.micadena.value == ',') {
	    	alert("Debe Seleccionar un Módulo.");
		}else{
	<% End If %>
		if (document.solpor.datossol.empleg.value=="") {
			alert("Debe ingresar un valor en Solicitado por");document.solpor.datossol.empleg.focus();
		}else{
		 if (document.relpor.datos.empleg.value=="") {
			alert("Debe ingresar un valor en Revelado por");document.relpor.datos.empleg.focus();
		}else{
		  if (!(validarfecha(document.datos.pedfecped))) {
		    document.datos.pedfecped.focus();
		  }else{
				if (!(validarfecha(document.datos.pedfecpret))) {
					document.datos.pedfecpret.focus();
				}else{
					if (!(menorque(document.datos.pedfecped.value,document.datos.pedfecpret.value))) {
						alert("La Fecha de Requerimiento debe ser menor o igual que la Fecha Pretendida.");document.datos.pedfecped.focus();}				
					else{
						if (isNaN(document.datos.peddurpredias.value)||(!(validanumero(document.datos.peddurpredias, 4, 0)))) {
					    	alert("La Duración en Días debe ser Numérica y Entera.");document.datos.peddurpredias.focus();
						}else{
							if (document.datos.peddurpredias.value < 0) {
					    		alert("La Duración en Días debe ser Positiva.");document.datos.peddurpredias.focus();
							 }else{
								if (isNaN(document.datos.peddurprethora.value)||(!(validanumero(document.datos.peddurprethora, 4, 0)))) {
					    			alert("La Cantidad de Horas por Clase debe ser Numérica y Entera.");document.datos.peddurprethora.focus();
								}else{
									if (document.datos.peddurprethora.value < 0) {
					    				alert("La Cantidad de Horas por Clase debe ser Positiva.");document.datos.peddurprethora.focus();
									}else{
										if (isNaN(document.datos.pedpers.value)||(!(validanumero(document.datos.pedpers, 4, 0)))) {
											alert("La Cantidad de Personas debe ser Numérica y Entera.");document.datos.pedpers.focus();
										}else{
											if (document.datos.pedpers.value < 0) {
					    						alert("La Cantidad de Personas debe ser Positiva.");document.datos.pedpers.focus();
											}else{
												if (isNaN(document.datos.pedprioridad.value)||(!(validanumero(document.datos.pedprioridad, 4, 0)))) {
													alert("La Prioridad debe ser Numérica y Entera.");document.datos.pedprioridad.focus();
												}else{
													if (document.datos.pedprioridad.value < 0) {
					    								alert("La Prioridad debe ser Positiva.");document.datos.pedprioridad.focus();
													}else{

									<% if l_HayAutorizacion then ' Si se debe tomar autorizacion ' firmas%>
										// Verifico que se haya cargado la autorización 
										if (((document.datos.seleccion.value == "") && (document.datos.seleccion1.value == "")) && ("<%= l_tipo %>" == "A"))
										    alert("Debe ingresar una autorización.");
										else{
									<% End If %>											
													
													
												
				<% If l_tipo = "A" then %>
					//alert(document.datos.micadena.value);
					abrirVentanaH('requerimientos_cap_03.asp?grabar=' + document.datos.micadena.value
					 + "&tipo=<%= l_tipo %>" 
					 //+ "&pednro=" + document.datos.pednro.value
					 + "&peddesabr=" + document.datos.peddesabr.value
					 + "&peddesext=" + document.datos.peddesext.value
					 + "&estpednro=2" //+ document.datos.estpednro.value     Por ahora siempre es 2 luego hay que agregar la seguridad
					//"&modnro=" + document.datos.modnro
					 + "&pedfecped=" + document.datos.pedfecped.value
					 + "&pedfecpret=" + document.datos.pedfecpret.value
					 + "&peddurpredias=" + document.datos.peddurpredias.value
					 + "&peddurprethora=" + document.datos.peddurprethora.value
					 + "&pedpers=" + document.datos.pedpers.value
					 + "&pedsolpor=" + document.solpor.datossol.empleg.value
					 + "&pedrelpor=" + document.relpor.datos.empleg.value
					 + "&pedprioridad=" + document.datos.pedprioridad.value
					 + "&pedmotprio=" + document.datos.pedmotprio.value
					 + "&seleccion=" + document.datos.seleccion.value
					 + "&seleccion1=" + document.datos.seleccion1.value , '','','');
					  }}
				 <% else %>
					 abrirVentanaH('requerimientos_cap_03.asp?'
					 + "&tipo=<%= l_tipo %>" 
					 + "&pednro=" + document.datos.pednro.value
					 + "&peddesabr=" + document.datos.peddesabr.value
					 + "&peddesext=" + document.datos.peddesext.value
					 + "&estpednro=" + document.datos.estpednro.value
					 + "&modnro=" + document.datos.modnro.value
					 + "&pedfecped=" + document.datos.pedfecped.value
					 + "&pedfecpret=" + document.datos.pedfecpret.value
					 + "&peddurpredias=" + document.datos.peddurpredias.value
					 + "&peddurprethora=" + document.datos.peddurprethora.value
					 + "&pedpers=" + document.datos.pedpers.value
					 + "&pedsolpor=" + document.solpor.datossol.empleg.value
					 + "&pedrelpor=" + document.relpor.datos.empleg.value
					 + "&pedprioridad=" + document.datos.pedprioridad.value
					 + "&pedmotprio=" + document.datos.pedmotprio.value 
 					 + "&seleccion=" + document.datos.seleccion.value
					 + "&seleccion1=" + document.datos.seleccion1.value , '','','');
				 <% End If %>
}}}}}}}}}}}}}}}
//'-----firmas-----
<% if l_HayAutorizacion then ' Si se debe tomar autorizacion %>
	}
<% End If %>
//'-----firmas-----									

function valido(){
  document.datos.submit();
}

function invalido(texto){
  alert(texto);
}

function Nuevo_Dialogo(w_in, pagina, ancho, alto)
{
 return w_in.showModalDialog(pagina,'', 'center:yes;dialogWidth:' + ancho.toString() + ';dialogHeight:' + alto.toString() + ';');
}
function Ayuda_Fecha(txt)
{
 var jsFecha = Nuevo_Dialogo(window, '/ticket/shared/js/calendar.html', 16, 15);

 if (jsFecha == null) txt.value = ''
 else txt.value = jsFecha;
}

function Firmas(){  // Para llamar a control de firmas, mandandole la descripcion y demas
	var path = "../gti/cysfirmas_00.asp?obj=document.all.seleccion&amp;tipo=<%= l_tipAutorizacion %>&amp;codigo=<%= l_temnro %>&amp;descripcion=requerimiento";
	abrirVentana(path,'_blank',421,180)
}
</script>

<body leftmargin="0" rightmargin="0" topmargin="0" bottommargin="0" onload="javascript:document.datos.peddesabr.focus()">
<form name="datos" action="contenidos_cap_03.asp?tipo=<%= l_tipo %>" method="post" >
<input type="Hidden" name="temnro" value="<%'= l_temnro %>">
<input type="Hidden" name="micadena" value="">
<input type="Hidden" name="seleccion" value="">
<input type="Hidden" name="seleccion1" value="">
<table cellspacing="0" cellpadding="0" border="0" width="100%" height="100%">
	<tr>
	    <td class="th2">Datos del Requerimiento</td>
		<td class="th2" align="right">		  
			<% if l_HayAutorizacion then ' Si se debe tomar autorizacion %>
			<% call MostrarBoton ("sidebtnSHW", "Javascript:Firmas();","Autorizar") %>
			&nbsp;&nbsp;&nbsp;
			<% End If %>		
			<a class=sidebtnHLP href="Javascript:ayuda('<%= Request.ServerVariables("SCRIPT_NAME")%>');">Ayuda</a>
		</td>
	</tr>
	<tr>
		<td colspan="2"></td>
	</tr>
	<tr >
		<td colspan="2" width="100" height="100%">
			<table cellspacing="0" cellpadding="0" border="0" width="100%" height="100%">
			<% If  l_tipo = "M" then %>	
				<tr>
				    <td align="right"><b>Código:</b></td>
					<td colspan="3">
						<input type="text" name="pednro" style="background : #e0e0de;" readonly size="4" maxlength="4" value="<%= l_pednro %>">
					</td>
				</tr>
			<% End If %>
				<tr>
				    <td align="right"><b>Desc. Abreviada:</b></td>
					<td colspan="3">
						<input type="text" name="peddesabr" style="width:380;"maxlength="50" value="<%= l_peddesabr %>">
					</td>
				</tr>
				
				<tr>
						<% if l_tipo = "A" then %> 
						<td align="right"><b>Curso:</b></td>
							<td colspan="3">
								<select style="width:383px" name=curnro size="1" onchange="document.datos.micadena.value='';document.ifrm.location = 'requisitomodulo_cap_00.asp?cabnro=' + document.datos.curnro.value">
								<option value=0 selected>«Seleccione una Opción»</option>
							<%	'Si es alta muestro los cursos, simo muestro los modulos
								Set l_rs = Server.CreateObject("ADODB.RecordSet")
								l_sql = "SELECT curnro, curdesabr"
								l_sql  = l_sql  & " FROM cap_curso "
								l_sql  = l_sql  & " ORDER BY curdesabr"
								rsOpen l_rs, cn, l_sql, 0
								do until l_rs.eof		%>	
								<option value= <%= l_rs("curnro") %> > 
								<%= l_rs("curdesabr") %> (<%=l_rs("curnro")%>) </option>
							<%			l_rs.Movenext
								loop
								l_rs.Close %>	
							</select>
							<script>// document.datos.curnro.value= "<%'= l_curnro %>"</script>
						<% else %>
						<td align="right"><b>Modulo:</b></td>
							<td colspan="3" align="left">
								<select style="width:383px" name=modnro size="1">
							<%	Set l_rs = Server.CreateObject("ADODB.RecordSet")
								l_sql = "SELECT modnro, moddesabr"
								l_sql  = l_sql  & " FROM cap_modulo "
								l_sql  = l_sql  & " ORDER BY moddesabr"
								rsOpen l_rs, cn, l_sql, 0
								do until l_rs.eof		%>	
								<option value= <%= l_rs("modnro") %> > 
								<%= l_rs("moddesabr") %> (<%=l_rs("modnro")%>) </option>
							<%			l_rs.Movenext
								loop
								l_rs.Close %>	
							</select>
							<script> document.datos.modnro.value= "<%= l_modnro %>"</script>
						<% end if %> 
					</td>	
				</tr>
				
				<tr>
					   <% If l_tipo = "A" then %>
				    <td align="center" colspan="4">			
							<iframe name="ifrm" src="requisitomodulo_cap_00.asp?cabnro=0" frameborder="0" marginheight="0" marginwidth="0" scrolling="no" style="zoom:1; background-color: #FAF0E6;" width="530" height="150">licho</iframe>
						<% Else  %>
					<td align="right"><b>Estado:</b></td>
					<td align="left" colspan="3">			
								<select style="width:383px" name=estpednro size="1" class="deshabinp"  disabled>
							<%	Set l_rs = Server.CreateObject("ADODB.RecordSet")
								l_sql = "SELECT estpednro, estpeddesabr"
								l_sql  = l_sql  & " FROM cap_estadopedido "
								l_sql  = l_sql  & " ORDER BY estpednro "
								rsOpen l_rs, cn, l_sql, 0
								do until l_rs.eof		%>	
								<option value= <%= l_rs("estpednro") %> > 
								<%= l_rs("estpeddesabr") %> (<%=l_rs("estpednro")%>) </option>
							<%			l_rs.Movenext
								loop
								l_rs.Close %>	
							</select>
							<script> document.datos.estpednro.value= "<%= l_estpednro %>"</script>
						<% End If %>
					</td>
				</tr>
				
				<tr>
				    <td align="right" nowrap width="0"><b>Fecha Requerimiento:</b></td>
					<td  align="left" nowrap width="0">
					    <input  type="Text" name="pedfecped" size="10" maxlength="10" value="<%= l_pedfecped %>">
						<a href="Javascript:Ayuda_Fecha(document.datos.pedfecped)"><img src="/ticket/shared/images/cal.gif" border="0"></a>	
					</td>
				    <td align="right" nowrap width="0"width="0"><b>Fecha Pretendida:</b></td>
					<td  align="left" nowrap width="0">
					    <input  type="Text" name="pedfecpret" size="10" maxlength="10" value="<%= l_pedfecpret %>">
						<a href="Javascript:Ayuda_Fecha(document.datos.pedfecpret)"><img src="/ticket/shared/images/cal.gif" border="0"></a>	
					</td>
				</tr>
				<tr>
				    <td align="right"><b>Solicitado por:</b></td>
					<td colspan="3" align="left">
					   <!-- <input  type="Text" name="estpednro" size="4" maxlength="4" value="<%'= l_pedsolpor %>">-->
						<iframe name="solpor" frameborder="0" width="100%" height="30" scrolling="No" src="requerimiento_sol_cap_00.asp?empleg=<%= l_pedsolpor %>"  ></iframe>
					</td>
				</tr>	
				<tr>
				    <td align="right"><b>Relevado por:</b></td>
					<td colspan="3" align="left">
					    <!--<input  type="Text" name="pedrelpor" size="4" maxlength="4" value="<%'= l_pedrelpor %>">-->
						<iframe name="relpor" frameborder="0" width="100%" height="30" scrolling="No" src="requerimiento_rel_cap_00.asp?empleg=<%= l_pedrelpor %>"  ></iframe>
					</td>
				</tr>
				<tr>
				    <td align="right"><b>Duración en Días:</b></td>
					<td  align="left">
					    <input  type="Text" name="peddurpredias" size="4" maxlength="4" value="<%= l_peddurpredias %>">
					</td>
				    <td align="right"><b>Horas por Clase:</b></td>
					<td  align="left">
					    <input  type="Text" name="peddurprethora" size="4" maxlength="4" value="<%= l_peddurprethora %>">
					</td>
				</tr>
				<tr>
					<td align="right"><b>Prioridad:</b></td>
					<td align="left">
					    <input  type="Text" name="pedprioridad" size="4" maxlength="4" value="<%= l_pedprioridad %>">
					</td>
				    <td align="right" nowrap><b>Cantidad Personas:</b></td>
					<td  align="left">
					    <input  type="Text" name="pedpers" size="4" maxlength="4" value="<%= l_pedpers %>">
					</td>
				</tr>
				<tr>
				    <td align="right"><b>Motivo de la Prioridad:</b></td>
					<td colspan="3" align="left">
					    <textarea name="pedmotprio" rows="3" cols="45" maxlength="200"><%= l_pedmotprio %></textarea>
					</td>
				</tr>
				<tr>
				    <td align="right"><b>Observaciones:</b></td>
					<td colspan="3" align="left">
					    <textarea name="peddesext" rows="3" cols="45" maxlength="200"><%= l_peddesext %></textarea>
					</td>
				</tr>
			</table>
		</td>
	</tr>		
	<tr>
	    <td  colspan="2" align="right" class="th2">
			<a class=sidebtnABM href="Javascript:Validar_Formulario()">Aceptar</a>
			<a class=sidebtnABM href="Javascript:window.close()">Cancelar</a>
		</td>
	</tr>
</table>
<!--<iframe name="valida" style="visibility=hidden;" src="" width="100%" height="100%"></iframe> -->
</form>
<%
set l_rs = nothing
'l_Cn.Close
'set l_Cn = nothing
%>
</body>
</html>

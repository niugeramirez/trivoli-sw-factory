<% Option Explicit %>
<!--#include virtual="/serviciolocal/shared/db/conn_db.inc"-->
<% 
'Archivo: embarque_con_02.asp
'Descripción: ABM Embarque
'Autor : Gustavo Manfrin
'Fecha: 18/09/2005
'Modificado: 

'Datos del formulario

'on error goto 0 

Dim l_embnro
Dim l_embcod
Dim l_embkiltot
Dim l_embkilemb
Dim l_embsim
Dim l_embact
Dim l_ordnro
Dim l_desnro
Dim l_desdes
Dim l_cornro
Dim l_entnro
Dim l_depnro
Dim l_connro

Dim l_empnro
Dim l_semnro

'ADO
Dim l_tipo
Dim l_sql
Dim l_rs

l_tipo = request.querystring("tipo")

%>
<html>
<head>
<link href="/serviciolocal/shared/css/tables3.css" rel="StyleSheet" type="text/css">
<meta http-equiv="Content-Type" content="text/html; charset=iso-8859-1">
<title><%= Session("Titulo")%>Embarque - Ticket</title>
</head>
<script src="/serviciolocal/shared/js/fn_windows.js"></script>
<script src="/serviciolocal/shared/js/fn_confirm.js"></script>
<script src="/serviciolocal/shared/js/fn_ayuda.js"></script>
<script src="/serviciolocal/shared/js/fn_valida.js"></script>
<script src="/serviciolocal/shared/js/fn_numeros.js"></script>
<script src="/serviciolocal/shared/js/fn_fechas.js"></script>
<script src="/serviciolocal/shared/js/fn_hora.js"></script>
<script src="/serviciolocal/shared/js/fn_ay_generica.js"></script>
<script>
function Validar_Formulario(){

if (document.datos.embcod.value== 0){

	alert("Debe ingresar el embarque.");
	document.datos.embcod.focus();
	return;
}

if (document.datos.embkiltot.value == 0){
	alert("Debe ingresar los kilos a embarcar.");
	document.datos.embkiltot.focus();
	return;
}

if (document.datos.ordnro.value == "0"){
	alert("Debe ingresar orden.");
	document.datos.ordnro.focus();
	return;
}

if (document.datos.cornro.value == ""){
	alert("Debe ingresar corredor.");
	document.datos.cornro.focus();
	return;
}

if (document.datos.desnro.value == ""){
	alert("Debe ingresar destinatario.");
	document.datos.desnro.focus();
	return;
}

if (document.datos.depnro.value == ""){
	alert("Debe ingresar deposito.");
	document.datos.depnro.focus();
	return;
}

var d=document.datos;
	
document.valida.location = "embarque_con_06.asp?tipo=<%= l_tipo%>&embnro="+document.datos.embnro.value+"&embcod="+document.datos.embcod.value+"&embact="+document.datos.embact.checked+"&ordnro="+document.datos.ordnro.value+"&connro="+document.datos.connro.value+"&desnro="+document.datos.desnro.value+"&cornro="+document.datos.cornro.value;

}

function valido(){
	document.datos.submit();
}

function invalido(texto){
	alert(texto);
	document.datos.embcod.focus();
}

function Obtener_Vencordes(){	
		if (isNaN(document.datos.desnro.value)) {
			alert("Debe ingresar un valor Númerico");
			document.datos.desnro.value = "";
			document.datos.desnro.focus();
			return;
		}
		if (document.datos.desnro.value == "") {
			document.datos.desnro.value = "";	
			document.datos.desdes.value = "NINGUNO";
		}
		else 
			document.valida.location = "embarque_con_07.asp?vencorcod=" + document.datos.desnro.value;	
}

function actualizar_vendedor(Vend){
	document.datos.desdes.value = Vend;

}

</script>
<% 
l_empnro = ""
l_semnro = ""	   
select Case l_tipo
	Case "A":
       l_embnro = ""
       l_embcod = ""
       l_embkiltot = ""
       l_embkilemb = 0
       l_embsim = ""
       l_embact = ""
	   l_desdes = "NINGUNO"
   	   l_desnro = ""
   	   l_cornro = ""
   	   l_depnro = ""
   	   l_entnro = ""
   	   l_ordnro = "0"
	   l_connro = "0"
       Set l_rs = Server.CreateObject("ADODB.RecordSet")
	   
 	Case "M":
		Set l_rs = Server.CreateObject("ADODB.RecordSet")
		l_embnro = request.querystring("cabnro")
		l_sql = "SELECT * "
		l_sql = l_sql & " FROM tkt_embarque "
		l_sql = l_sql & " LEFT JOIN tkt_vencor ON tkt_vencor.vencornro = tkt_embarque.desnro "		
      	l_sql = l_sql & " WHERE embnro = " & l_embnro
		rsOpen l_rs, cn, l_sql, 0 
		if not l_rs.eof then
           l_embcod = l_rs("embcod")
           l_embkiltot = l_rs("embkiltot")
           l_embkilemb = l_rs("embkilemb")		   
           l_embsim = l_rs("embsim")
           l_embact = l_rs("embact")
           l_cornro = l_rs("cornro")
           l_desnro = l_rs("vencorcod")
           l_entnro = l_rs("entnro")
           l_ordnro = l_rs("ordnro")
           l_depnro = l_rs("depnro")
		   l_desdes = l_rs("vencordes")
		   if isnull(l_rs("connro")) then
              l_connro = 0
		   else
	           l_connro = l_rs("connro")
		   end if 
		end if
		l_rs.Close
end select

'response.write l_connro & "-" & l_ordnro & "-"

%>
<body leftmargin="0" rightmargin="0" topmargin="0" bottommargin="0" onload="JavaScript:document.datos.pronro.focus()">
<form name="datos" action="embarque_con_03.asp?tipo=<%= l_tipo %>&embnro=<%= l_embnro %>&embcod=<%= l_embcod %>" method="post" target="valida">
<input type="Hidden" name="embnro" value="<%= l_embnro %>">

<table cellspacing="0" cellpadding="0" border="0" width="100%" height="100%">
<tr>
    <td class="th2" nowrap>Embarque</td>
	<td class="th2" align="right">
		<a class=sidebtnHLP href="Javascript:ayuda('<%= Request.ServerVariables("SCRIPT_NAME")%>');">Ayuda</a>
	</td>
</tr>
<tr>
	<td colspan="2" height="100%">
		<table border="0" cellspacing="0" cellpadding="0">
			<tr>
				<td width="50%"></td>
				<td>
					<table cellspacing="0" cellpadding="0" >
					<tr>
						<td align="right" nowrap><b>Embarque:</b></td>
						<td>
							<input type="text" name="embcod" size="9" maxlength="8" value="<%= l_embcod %>">
						</td>
					</tr>
					<tr>
						<td align="right" nowrap><b>Kilos Embarque:</b></td>					
						<td>
							<input type="text" name="embkiltot" size="9" maxlength="8" value="<%= l_embkiltot %>">
						</td>
					</tr>
					<tr>
						<td align="right" nowrap><b>Kilos Embarcados:</b></td>					
						<td>
							<input type="text" readonly class="deshabinp" name="embkilemb" size="9" maxlength="8" value="<%= l_embkilemb %>">
						</td>
					</tr>
					<tr>
					    <td align="right" nowrap><b>Orden Trabajo:</b></td>
						<td colspan="3">
							<select name="ordnro" size="1" style="width:355;" >
							<option value=0 selected>&laquo; Seleccione una orden  &raquo;</option>
							<%	l_sql = "SELECT ordnro, ordcod, prodes, procod "
 							    l_sql  = l_sql  & " FROM tkt_ordentrabajo "
								l_sql  = l_sql  & " INNER JOIN tkt_producto ON tkt_ordentrabajo.pronro = tkt_producto.pronro"
	  						    l_sql  = l_sql  & " WHERE (ordhab = -1) "								
  							    l_sql  = l_sql  & " ORDER BY ordcod "
								rsOpen l_rs, cn, l_sql, 0
								do until l_rs.eof %>	
								<option value=<%= l_rs("ordnro") %> > 
								<%= l_rs("ordcod") & " - (" & l_rs("procod") & ") - "  &  l_rs("prodes")%> </option>
								<%	l_rs.Movenext
								loop
								l_rs.Close %>
							</select>
							<script> document.datos.ordnro.value= "<%= l_ordnro %>"</script>
							</td>
					 </tr>

					<tr>
					    <td align="right" nowrap><b>Contrato:</b></td>
						<td colspan="3">
							<select name="connro" size="1" style="width:355;" >
							<option value=0 selected>&laquo; Ninguno &raquo;</option>
							<%	l_sql = "SELECT connro, concod, prodes, vencordes, procod "
 							    l_sql  = l_sql  & " FROM tkt_contrato "
								l_sql  = l_sql  & " INNER JOIN tkt_producto ON tkt_contrato.pronro = tkt_producto.pronro"
								l_sql  = l_sql  & " INNER JOIN tkt_vencor ON tkt_contrato.vennro = tkt_vencor.vencornro"								
 							    l_sql  = l_sql  & " WHERE (conact = -1) "
								l_sql  = l_sql  & " ORDER BY concod "
								rsOpen l_rs, cn, l_sql, 0
								do until l_rs.eof %>	
								<option value=<%= l_rs("connro") %> > 
								<%= l_rs("concod") & " - (" &  l_rs("procod") & ") - " &  l_rs("vencordes") %> </option>
								<%	l_rs.Movenext
								loop
								l_rs.Close %>
							</select>
							<script> document.datos.connro.value= "<%= l_connro %>"</script>
							</td>
					 </tr>

	 				 <tr>
						    <td align="right" nowrap><b>Corredor:</b></td>
							<td colspan="3">
								<select name="cornro" size="1" style="width:355;" >
									<option value=0 selected>&laquo; Seleccione un Corredor &raquo;</option>
								<%	l_sql = "SELECT vencornro, vencordes, vencorcod "
									l_sql  = l_sql  & " FROM tkt_vencor WHERE vencortip='C' "
									rsOpen l_rs, cn, l_sql, 0
									do until l_rs.eof %>	
									<option value=<%= l_rs("vencornro") %> > 
									<%= l_rs("vencordes") %> (<%=l_rs("vencorcod")%>) </option>
									<%	l_rs.Movenext
									loop
									l_rs.Close %>
								</select>
								<script> document.datos.cornro.value="<%= l_cornro %>"</script>
							</td>
					</tr>										

					<tr>
						<td align="right" nowrap><b>Destinatario:</b></td>
						<td align="left" nowrap>
							<input " onblur="javascript:Obtener_Vencordes();" type="text" size="7" name="desnro"  value="<%=l_desnro %>"   >&nbsp;
							<button onclick="javascript:ayudacodigo(document.datos.desnro,document.datos.desdes,'vencorcod','vencordes','tkt_vencor','vencortip=\'V\' and venact=-1','Código;Descripción','Destinatarios');"> ^ </button>
							<input type="text" name="desdes" size="40" maxlength="45" value="<%=l_desdes %>" readonly>		
						</td>
					</tr>				
					<tr>
						<td align="right" nowrap><b>Depósito:</b></td>
						<td>
						<select name="depnro" size="1" style="width:355;" <%'= l_claseCombo %>" >
						<option value="" selected>&laquo; Seleccione un Depósito &raquo;</option>
						<%	l_sql = "SELECT depnro, depdes, depcod "
							l_sql  = l_sql  & " FROM tkt_deposito "
							l_sql  = l_sql  & " ORDER BY depdes "
							rsOpen l_rs, cn, l_sql, 0
							do until l_rs.eof %>	
								<option value=<%= l_rs("depnro") %> > 
								<%= l_rs("depdes") %> (<%=l_rs("depcod")%>) </option>
								<%	l_rs.Movenext
							loop
							l_rs.Close %>
						</select>
						<script> document.datos.depnro.value= "<%= l_depnro %>"</script>
						</td>
					</tr>
					<tr>
					    <td align="right"><b>S.I.M:</b></td>
						<td>
							<TEXTAREA name="embsim" rows="3" cols="42" ><%= l_embsim %></TEXTAREA>
						</td>
					</tr>
				     <tr>
  					    <td height="100%" align="right" nowrap><b>Activo:</b></td>
	 					<td height="100%">
   							<input type="Checkbox" name="embact" <% If l_embact = -1 then  %>checked<% end if %> 
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
    <td colspan="2" align="right" class="th2">
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

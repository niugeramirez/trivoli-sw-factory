
<% Option Explicit %>
<!--#include virtual="/serviciolocal/shared/db/conn_db.inc"-->
<% 
'Archivo: contracts_con_02.asp
'Descripción: ABM de Contracts
'Autor : Raul Chinestra
'Fecha: 27/11/2007

on error goto 0

'Datos del formulario

Dim l_legpar1
Dim l_legpar2
Dim l_legpar3
Dim l_serdes
Dim l_legnro
Dim l_legape
Dim l_legnom
Dim l_legfecing
Dim l_legdni
Dim l_legfecnac
Dim l_legdom
Dim l_legtel
Dim l_pronro
Dim l_legapenommad
Dim	l_legdommad
Dim	l_legtelmad
Dim l_legapenompad
Dim	l_legdompad
Dim	l_legtelpad
Dim l_legins
Dim l_leginsedu
Dim l_legcobsoc
Dim l_legabo
Dim l_mednro

Dim l_buqnro
Dim l_buqdes
Dim l_tipopenro
Dim l_tipbuqnro
Dim l_agenro
Dim l_buqfecdes
Dim l_buqfechas
Dim l_buqton

'ADO
Dim l_tipo
Dim l_sql
Dim l_rs

l_tipo = request.querystring("tipo")

'response.write l_tipo

%>
<html>
<head>
<link href="/serviciolocal/ess/shared/css/tables_gray.css" rel="StyleSheet" type="text/css">
<!--<link href="/serviciolocal/shared/css/tables_gray.css" rel="StyleSheet" type="text/css">-->
<meta http-equiv="Content-Type" content="text/html; charset=iso-8859-1">
<title><%= Session("Titulo")%>Servicio Local - Legajo</title>
</head>
<script src="/serviciolocal/shared/js/fn_valida.js"></script>
<script src="/serviciolocal/shared/js/fn_fechas.js"></script>
<script src="/serviciolocal/shared/js/fn_ayuda.js"></script>
<script src="/serviciolocal/shared/js/fn_windows.js"></script>
<script src="/serviciolocal/shared/js/fn_numeros.js"></script>
<script>
function Validar_Formulario(){
/*
if (document.datos.buqdes.value == ""){
	alert("Debe ingresar el Nombre del Buque.");
	document.datos.buqdes.focus();
	return;
}

if (document.datos.tipopenro.value == 0){
	alert("Debe ingresar el Tipo de Operación.");
	document.datos.tipopenro.focus();
	return;
}
if (document.datos.tipbuqnro.value == 0){
	alert("Debe ingresar el Tipo de Buque.");
	document.datos.tipbuqnro.focus();
	return;
}
if (document.datos.agenro.value == 0){
	alert("Debe ingresar la Agencia.");
	document.datos.agenro.focus();
	return;
}

if ((document.datos.buqfecdes.value != "")&&(!validarfecha(document.datos.buqfecdes))){
	 document.datos.buqfecdes.focus();
	 return;
}

if ((document.datos.buqfechas.value != "")&&(!validarfecha(document.datos.buqfechas))){
	 document.datos.buqfechas.focus();
	 return;
}

if ((document.datos.buqfecdes.value != "")&&(document.datos.buqfechas.value != "") ){

	if (!(menorque(document.datos.buqfecdes.value,document.datos.buqfechas.value))) {
			alert("La Fecha de Comienzo debe ser menor o igual que la Fecha de Termino.");
			document.datos.buqfecdes.focus();
		    return;			
	}		
}	
*/

/*
var d=document.datos;
document.valida.location = "countries_con_06.asp?tipo=<%= l_tipo%>&counro="+document.datos.counro.value + "&coudes="+document.datos.coudes.value;
*/
valido();
}

function valido(){
	document.datos.submit();
}

function invalido(texto){
	alert(texto);
	document.datos.coudes.focus();
}

function actualizaBerth(valor){

   document.datos.bernro.value = 0 ;    

  if ((document.datos.pornro.value == "")||(document.datos.pornro.value == "0"))   	 
 	 document.ifrmBerth.location = "contracts_berth_con_00.asp?pornro=0&disabled=disabled";
  else 
     document.ifrmberth.location = "contracts_berth_con_00.asp?pornro="+valor+"&bernro=0";  
}

function Nuevo_Dialogo(w_in, pagina, ancho, alto)
{
 return w_in.showModalDialog(pagina,'', 'center:yes;dialogWidth:' + ancho.toString() + ';dialogHeight:' + alto.toString() + ';');
}
function Ayuda_Fecha(txt)
{
 var jsFecha = Nuevo_Dialogo(window, '/serviciolocal/shared/js/calendar.html', 16, 15);

 if (jsFecha == null) txt.value = ''
 else txt.value = jsFecha;
}

</script>
<% 
select Case l_tipo
	Case "A":
    	l_legpar1      = ""
    	l_legpar2      = ""
    	l_legpar3      = ""
		l_serdes       = ""
		l_legfecing	   = ""
		l_legape	   = ""
		l_legnom       = ""
		l_legdni       = ""
		l_legfecnac    = ""
		l_legdom       = ""
		l_legtel       = ""
		l_pronro       = 0
		l_legapenommad = ""
		l_legdommad    = ""
		l_legtelmad    = ""
		l_legins	   = ""
		l_leginsedu    = ""
		l_legcobsoc    = ""
		l_legabo	   = ""
		l_mednro       = 0
	Case "M":
		Set l_rs = Server.CreateObject("ADODB.RecordSet")
		l_legnro = request.querystring("cabnro")
		l_sql = "SELECT  * "
		l_sql = l_sql & " FROM ser_legajo "
		l_sql = l_sql & " INNER JOIN ser_servicio ON ser_servicio.sercod = ser_legajo.legpar1 "
		l_sql  = l_sql  & " WHERE legnro = " & l_legnro
		rsOpen l_rs, cn, l_sql, 0 
		if not l_rs.eof then
	    	l_legpar1      = l_rs("legpar1")
	    	l_legpar2      = l_rs("legpar2")
	    	l_legpar3      = l_rs("legpar3")
	    	l_serdes       = l_rs("serdes")
			l_legfecing	   = l_rs("legfecing")
			l_legape	   = l_rs("legape")
			l_legnom       = l_rs("legnom")
			l_legdni       = l_rs("legdni")
			l_legfecnac	   = l_rs("legfecnac")
			l_legdom       = l_rs("legdom")
			l_legtel	   = l_rs("legtel")
			l_pronro       = l_rs("pronro")
			l_legapenommad = l_rs("legapenommad")
			l_legdommad    = l_rs("legdommad")
			l_legtelmad    = l_rs("legtelmad")
			l_legins	   = l_rs("legins")
			l_leginsedu    = l_rs("leginsedu")
			l_legcobsoc    = l_rs("legcobsoc")
			l_legabo	   = l_rs("legabo")
			l_mednro       = l_rs("mednro")
		end if
		l_rs.Close
end select

%>
<body leftmargin="0" rightmargin="0" topmargin="0" bottommargin="0" onload="javascript:document.datos.legpar1.focus();">
<form name="datos" action="legajos_con_03.asp?tipo=<%= l_tipo %>" method="post" target="valida">
<input type="Hidden" name="legnro" value="<%= l_legnro %>">

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
						<td align="left" colspan="4" style="font-size:20"  >
							Servicio Local: <b><%= l_serdes %><b>				
						</td>																	
					</tr>							
					<tr>
					    <td align="right" ><b>Legajo:</b></td>
						<td align="left" colspan="3"  >
						    <input type="text" name="legpar1" size="2" maxlength="2" value="<%= l_legpar1 %>">
						    <input type="text" name="legpar2" size="10" maxlength="10" value="<%= l_legpar2 %>">
						    <input type="text" name="legpar3" size="2" maxlength="2" value="<%= l_legpar3 %>">							
						</td>																	
					</tr>																
					<tr>
					    <td align="right" ><b>Fecha Ingreso:</b></td>
						<td align="left" colspan="3"  >
						    <input type="text" name="legfecing" size="10" maxlength="10" value="<%= l_legfecing %>">
							<a href="Javascript:Ayuda_Fecha(document.datos.legfecing)"><img src="/serviciolocal/shared/images/cal.gif" border="0"></a>
						</td>																	
					</tr>											
					<tr>
					    <td align="right"><b>Apellido:</b></td>
						<td>
							<input type="text" name="legape" size="20" maxlength="20" value="<%= l_legape %>">							
						</td>
					    <td align="right"><b>Nombre:</b></td>						
						<td>
							<input type="text" name="legnom" size="20" maxlength="20" value="<%= l_legnom %>">
						</td>						
					</tr>					
					<tr>
					    <td align="right"><b>D.N.I.:</b></td>
						<td>
							<input type="text" name="legdni" size="20" maxlength="20" value="<%= l_legdni %>">
						</td>
					    <td align="right" ><b>Fec. Nac.:</b></td>
						<td align="left"  >
						    <input type="text" name="legfecnac" size="10" maxlength="10" value="<%= l_legfecnac %>">
							<a href="Javascript:Ayuda_Fecha(document.datos.legfecnac)"><img src="/serviciolocal/shared/images/cal.gif" border="0"></a>
						</td>						
					</tr>
					<tr>
					    <td align="right"><b>Domicilio:</b></td>
						<td>
							<input type="text" name="legdom" size="20" maxlength="20" value="<%= l_legdom %>">
						</td>
						<td align="right"><b>Teléfono:</b></td>
						<td>
							<input type="text" name="legtel" size="20" maxlength="20" value="<%= l_legtel %>">
						</td>						
					</tr>

					<tr>
						<td  align="right" nowrap><b>Derecho Vulnerado: </b></td>
						<td colspan="3"><select name="pronro" size="1" style="width:150;">
								<option value=0 selected>Todos</option>
								<%Set l_rs = Server.CreateObject("ADODB.RecordSet")
								l_sql = "SELECT  * "
								l_sql  = l_sql  & " FROM ser_problematica "
								l_sql  = l_sql  & " ORDER BY prodes "
								rsOpen l_rs, cn, l_sql, 0
								do until l_rs.eof		%>	
								<option value= <%= l_rs("pronro") %> > 
								<%= l_rs("prodes") %> (<%=l_rs("pronro")%>) </option>
								<%	l_rs.Movenext
								loop
								l_rs.Close %>
							</select>
							<script>document.datos.pronro.value= "<%= l_pronro %>"</script>
						</td>					
					</tr>
					<tr>
					    <td align="right"><b>Madre - Apellido y Nombre:</b></td>
						<td>
							<input type="text" name="legapenommad" size="20" maxlength="20" value="<%= l_legapenommad %>">
						</td>
						<td align="right"><b>Dom:</b></td>						
						<td>
							<input type="text" name="legdommad" size="20" maxlength="20" value="<%= l_legdommad %>">
							<b>Tel:</b> <input type="text" name="legtelmad" size="10" maxlength="20" value="<%= l_legtelmad %>">						
						</td>							
					</tr>																				
					<tr>
					    <td align="right"><b>Padre - Apellido y Nombre:</b></td>
						<td>
							<input type="text" name="legapenompad" size="20" maxlength="20" value="<%= l_legapenompad  %>">
						</td>
						<td align="right"><b>Dom:</b></td>												
						<td>
							<input type="text" name="legdompad" size="20" maxlength="20" value="<%= l_legdompad %>">
							<b>Tel:</b> <input type="text" name="legtelpad" size="10" maxlength="20" value="<%= l_legtelpad %>">
						</td>						
					</tr>					
					<tr>
					    <td align="right"><b>Instituciones Intervinientes:</b></td>
						<td colspan="3">
							<input type="text" name="legins" size="80" maxlength="20" value="<%= l_legins %>">
						</td>
					</tr>																				
					<tr>
					    <td align="right"><b>Instituciones Educativas:</b></td>
						<td colspan="3">
							<input type="text" name="leginsedu" size="80" maxlength="20" value="<%= l_leginsedu %>">
						</td>
					</tr>																									
					<tr>
					    <td align="right"><b>Cobertura Social de la Familia:</b></td>
						<td colspan="3">
							<input type="text" name="legcobsoc" size="80" maxlength="20" value="<%= l_legcobsoc %>">
						</td>
					</tr>																														
					<tr>
					    <td align="right"><b>Estrategias de Intervención:</b></td>
						<td colspan="3">
							<input type="text" name="legabo" size="80" maxlength="20" value="<%= l_legabo %>">
						</td>
					</tr>					
					<tr>
						<td align="right"><b>Medidas Protección:</b></td>
						<td colspan="3"><select name="mednro" size="1" style="width:150;">
								<option value=0 selected>&nbsp;</option>
								<%Set l_rs = Server.CreateObject("ADODB.RecordSet")
								l_sql = "SELECT  * "
								l_sql  = l_sql  & " FROM ser_medida "
								l_sql  = l_sql  & " ORDER BY meddes "
								rsOpen l_rs, cn, l_sql, 0
								do until l_rs.eof		%>	
								<option value= <%= l_rs("mednro") %> > 
								<%= l_rs("meddes") %> (<%=l_rs("mednro")%>) </option>
								<%	l_rs.Movenext
								loop
								l_rs.Close %>
							</select>
							<script>document.datos.mednro.value= "<%= l_mednro %>"</script>
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
<%
set l_rs = nothing
cn.Close
set cn = nothing
%>
</body>
</html>

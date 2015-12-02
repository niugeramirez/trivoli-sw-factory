
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
Dim l_descripcion
'ADO
Dim l_tipo
Dim l_sql
Dim l_rs

Dim l_hd
Dim l_md
Dim l_hh
Dim l_mh

Dim l_fecha
Dim l_idrecursoreservable

l_tipo = request.querystring("tipo")
l_id = request("cabnro")


%>
<html>
<head>
<link href="/turnos/ess/shared/css/tables_gray.css" rel="StyleSheet" type="text/css">
<meta http-equiv="Content-Type" content="text/html; charset=iso-8859-1">
<title>Asignar Turnos</title>
</head>
<script src="/turnos/shared/js/fn_valida.js"></script>
<script src="/turnos/shared/js/fn_fechas.js"></script>
<script src="/turnos/shared/js/fn_ayuda.js"></script>
<script src="/turnos/shared/js/fn_windows.js"></script>
<script src="/turnos/shared/js/fn_numeros.js"></script>

<!-- Comienzo Datepicker -->
<link rel="stylesheet" href="../js/themes/smoothness/jquery-ui.css">
<script src="../js/jquery-1.8.0.js"></script>
<script src="../js/jquery-ui.js"></script>  
<script src="../js/jquery.ui.datepicker-es.js"></script>
<script>
$(function () {
$.datepicker.setDefaults($.datepicker.regional["es"]);
$("#datepicker").datepicker({
firstDay: 1
});

		
$( "#fechadesde" ).datepicker({
	showOn: "button",
	buttonImage: "/turnos/shared/images/calendar1.png",
	buttonImageOnly: true
});

$( "#fechahasta" ).datepicker({
	showOn: "button",
	buttonImage: "/turnos/shared/images/calendar1.png",
	buttonImageOnly: true
});

});
</script>
<!-- Final Datepicker -->


<script>
function Validar_Formulario(){

if (document.datos.fechadesde.value == "")  {
 		alert("Debe ingresar la Fecha Desde ");
 		document.datos.fechadesde.focus();
	return;
}

if (document.datos.fechahasta.value == "")  {
 		alert("Debe ingresar la Fecha Hasta ");
 		document.datos.fechahasta.focus();
	return;
}	

if (!menorque(document.datos.fechadesde.value ,document.datos.fechahasta.value)){
	alert('La fecha hasta debe ser mayor que la fecha desde.');
 		document.datos.fechadesde.focus();
	return;
}	

if (document.datos.motivo.value == ""){
	alert("Debe ingresar el Motivo.");
	document.datos.motivo.focus();
	return;
}

/*
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


var d=document.datos;


document.valida.location = "anularTurno_con_06.asp?tipo=<%= l_tipo%>&id="+document.datos.id.value + "&qfechadesde=" + document.all.fechadesde.value + "&qfechahasta=" + document.all.fechahasta.value + "&opc="+document.datos.rbopc.value + "&hd="+document.datos.hd.value + "&md="+document.datos.md.value+ "&hh="+document.datos.hh.value+ "&mh="+document.datos.mh.value + "&idrecursoreservable="+document.datos.idrecursoreservable.value;

//valido();
}

function Habilitar(opc){

switch (opc.value) {
   case '1' :
      document.datos.hd.disabled  = true;
	  document.datos.md.disabled  = true;
      document.datos.hh.disabled  = true;
	  document.datos.mh.disabled  = true;		  
	  document.datos.fechadesde.disabled  = true;
	  document.datos.fechahasta.disabled  = true;
	  
	  document.datos.hd.readOnly = true;  
	  document.datos.hd.className="deshabinp"
	  
	  document.datos.md.readOnly = true;  
	  document.datos.md.className="deshabinp"
	  
	  document.datos.hh.readOnly = true;  
	  document.datos.hh.className="deshabinp"
	  
	  document.datos.mh.readOnly = true;  
	  document.datos.mh.className="deshabinp"	  	  	 
	  
	  document.datos.fechadesde.readOnly = true;  
	  document.datos.fechadesde.className="deshabinp"		
	  
	  document.datos.fechahasta.readOnly = true;  
	  document.datos.fechahasta.className="deshabinp"	
	  
     
	    	  
	  break;
   case '2' :
      document.datos.hd.disabled  = false;
	  document.datos.md.disabled  = false;
      document.datos.hh.disabled  = false;
	  document.datos.mh.disabled  = false;	
	  document.datos.fechadesde.disabled  = false;
	  document.datos.fechahasta.disabled  = false;	  
	  
	  document.datos.hd.readOnly = false;  
	  document.datos.hd.className="habinp"
	  
	  document.datos.md.readOnly = false;  
	  document.datos.md.className="habinp"
	  
	  document.datos.hh.readOnly = false;  
	  document.datos.hh.className="habinp"
	  
	  document.datos.mh.readOnly = false;  
	  document.datos.mh.className="habinp"	
	  
	  document.datos.fechadesde.readOnly = false;  
	  document.datos.fechadesde.className="habinp"		
	  
	  document.datos.fechahasta.readOnly = false;  
	  document.datos.fechahasta.className="habinp"		  
	  
	  break;
} 

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
 var jsFecha = Nuevo_Dialogo(window, '/turnos/shared/js/calendar.html', 16, 15);

 if (jsFecha == null) txt.value = ''
 else txt.value = jsFecha;
}

</script>
<% 

Set l_rs = Server.CreateObject("ADODB.RecordSet")
'l_id = request.querystring("cabnro")
l_sql = "SELECT fechahorainicio,  CONVERT(VARCHAR(10), fechahorainicio, 101) AS DateOnly , idrecursoreservable, descripcion "
l_sql = l_sql & " FROM calendarios "
l_sql = l_sql & " INNER JOIN recursosreservables ON recursosreservables.id = calendarios.idrecursoreservable "
l_sql  = l_sql  & " WHERE calendarios.id = " & l_id
rsOpen l_rs, cn, l_sql, 0 
if not l_rs.eof then
   	l_fecha      = left(l_rs("fechahorainicio"),10)
	l_idrecursoreservable = l_rs("idrecursoreservable")
	l_descripcion = l_rs("descripcion")
end if
l_rs.Close


%>
<body leftmargin="0" rightmargin="0" topmargin="0" bottommargin="0" onload="javascript:document.datos.motivo.focus();">
<form name="datos" action="AnularTurnos_con_03.asp?tipo=<%= l_tipo %>" method="post" target="valida">
<input type="hidden" name="id" value="<%= l_id %>">
<input type="hidden" name="idrecursoreservable" value="<%= l_idrecursoreservable %>">

<table cellspacing="0" cellpadding="0" border="0" width="100%" height="100%">
<tr>
    <td class="th2" nowrap>Medico: <%= l_descripcion %></td>
	<td class="th2" align="right">
		&nbsp;
	</td>
</tr>
<tr>
    <td class="th2" nowrap><% if l_tipo = "B" then response.write "Bloquear un Calendario o Rango de Calendarios" else response.write "Desbloquear un Calendario o Rango de Calendarios" end if%></td>
	<td class="th2" align="right">
		&nbsp;
	</td>
</tr>
<tr>
	<td colspan="2" height="100%">
		<table border="0" cellspacing="0" cellpadding="0">
			<tr>
				<td>
					<table cellspacing="0" cellpadding="0" border="0">
					
					 <tr> 
					  	<td colspan="4" align="center"> 	
								<input type=radio name=rbopc value=1 CHECKED onclick="Habilitar(this)"> <b>Calendario Actual</b>
					   			<input type=radio name=rbopc value=2 onclick="Habilitar(this)"> <b>Rango de Calendarios</b>	</td>
					  </tr>							
					 
					<tr>
						<td align="right" nowrap><b>Fecha Desde: </b></td>
						<td align="left"><input disabled class="deshabinp" id="fechadesde" type="text" name="fechadesde" value="<%= l_fecha %>"></td>
						
						
						<td align="right" nowrap><b>Fecha Hasta: </b></td>
						<td align="left"><input disabled class="deshabinp" id="fechahasta" type="text" name="fechahasta" value="<%= l_fecha %>"></td>	
																				
					</tr>  
					
					 <tr> 
					  	<td colspan="4" align="center">	&nbsp;	</td>
					 </tr>								

					  
					 <tr> 
					  	<td colspan="4" align="center">	&nbsp;	</td>
					 </tr>					  				
						
					<tr>
						<td  align="right" nowrap><b>Desde: </b></td>
						<td ><select disabled class="deshabinp" name="hd" size="1" style="width:50;">
								<%
								l_hd = 0  
								do while clng(l_hd) < 24 %>
								<option value= <%= right("0" & l_hd, 2) %>> <%= right("0" & l_hd, 2) %> </option>
								<%	l_hd = clng(l_hd) + 1
								loop
								%>
							</select>							
						    <b>:</b>
							<select disabled class="deshabinp" name="md" size="1" style="width:50;">
								<%
								l_md = 0  
								do while clng(l_md) < 60 %>
								<option value= <%= right("0" & l_md, 2) %>> <%= right("0" & l_md, 2) %> </option>
								<%	l_md = clng(l_md) + 15
								loop
								%>
							</select>							
						</td>			
						<td  align="right" nowrap><b>Hasta: </b></td>
						<td ><select disabled class="deshabinp" name="hh" size="1" style="width:50;">
								<%
								l_hh = 0  
								do while clng(l_hh) < 24 %>
								<option value= <%= right("0" & l_hh, 2) %>> <%= right("0" & l_hh, 2) %> </option>
								<%	l_hh = clng(l_hh) + 1
								loop
								%>
							</select>	
							<script>document.datos.hh.value="23"</script>							
						<b>:</b>
						   <select disabled class="deshabinp" name="mh" size="1" style="width:50;">
								<%
								l_mh = 0  
								do while clng(l_mh) < 60 %>
								<option value= <%= right("0" & l_mh, 2) %>> <%= right("0" & l_mh, 2) %> </option>
								<%	l_mh = clng(l_mh) + 15
								loop
								%>
							</select>							
						</td>						    																
					</tr>  	
					<!--															
					<tr>
					    <td align="right" ><b>Fecha Ingreso:</b></td>
						<td align="left" colspan="3"  >
						    <input type="text" name="legfecing" size="10" maxlength="10" value="<%'= l_legfecing %>">
							<a href="Javascript:Ayuda_Fecha(document.datos.legfecing)"><img src="/turnos/shared/images/cal.gif" border="0"></a>
						</td>																	
					</tr>	-->			
					 <tr> 
					  	<td colspan="4" align="center">	&nbsp;	</td>
					 </tr>
					 <tr> 
					  	<td colspan="4" align="center">	&nbsp;	</td>
					 </tr>					 												
					<tr>
					    <td align="right"><b>Motivo:</b></td>
						<td colspan="3">
							<input type="text" name="motivo" size="87" maxlength="50" value="<%'= l_apellido %>">							
						</td>
					  
												
					</tr>					
					
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

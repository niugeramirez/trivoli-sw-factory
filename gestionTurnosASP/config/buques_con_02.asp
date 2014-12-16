
<% Option Explicit %>
<!--#include virtual="/serviciolocal/shared/db/conn_db.inc"-->
<% 
'Archivo: contracts_con_02.asp
'Descripción: ABM de Contracts
'Autor : Raul Chinestra
'Fecha: 27/11/2007

on error goto 0

'Datos del formulario
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

%>
<html>
<head>
<link href="/serviciolocal/ess/shared/css/tables_gray.css" rel="StyleSheet" type="text/css">
<!--<link href="/serviciolocal/shared/css/tables_gray.css" rel="StyleSheet" type="text/css">-->
<meta http-equiv="Content-Type" content="text/html; charset=iso-8859-1">
<title><%= Session("Titulo")%> Buques - Buques</title>
</head>
<script src="/serviciolocal/shared/js/fn_valida.js"></script>
<script src="/serviciolocal/shared/js/fn_fechas.js"></script>
<script src="/serviciolocal/shared/js/fn_ayuda.js"></script>
<script src="/serviciolocal/shared/js/fn_windows.js"></script>
<script src="/serviciolocal/shared/js/fn_numeros.js"></script>
<script>
function Validar_Formulario(){

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
		l_buqdes	   = ""
		l_tipopenro    = 0
		l_tipbuqnro    = 0
		l_agenro       = 0
		l_buqfecdes    = ""
		l_buqfechas    = ""
		l_buqton	   = 0
	Case "M":
		Set l_rs = Server.CreateObject("ADODB.RecordSet")
		l_buqnro = request.querystring("cabnro")
		l_sql = "SELECT  * "
		l_sql = l_sql & " FROM buq_buque "
		'l_sql = l_sql & " INNER JOIN for_area ON for_area.arenro = for_country.arenro "
		l_sql  = l_sql  & " WHERE buqnro = " & l_buqnro
		rsOpen l_rs, cn, l_sql, 0 
		if not l_rs.eof then
			l_buqdes	   = l_rs("buqdes")	
			l_tipopenro    = l_rs("tipopenro")
			l_tipbuqnro    = l_rs("tipbuqnro")
			l_agenro       = l_rs("agenro")
			l_buqfecdes    = l_rs("buqfecdes")
			l_buqfechas    = l_rs("buqfechas")
			l_buqton	   = l_rs("buqton")
		end if
		l_rs.Close
end select

%>
<body leftmargin="0" rightmargin="0" topmargin="0" bottommargin="0" onload="javascript:document.datos.buqdes.focus();">
<form name="datos" action="buques_con_03.asp?tipo=<%= l_tipo %>" method="post" target="valida">
<input type="Hidden" name="buqnro" value="<%= l_buqnro %>">

<table cellspacing="0" cellpadding="0" border="0" width="100%" height="100%">
<tr>
    <td class="th2" nowrap>Buques</td>
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
					    <td align="right"><b>Nombre:</b></td>
						<td>
							<input type="text" name="buqdes" size="20" maxlength="20" value="<%= l_buqdes %>">
						</td>
					</tr>
					<tr>
						<td align="right"><b>Tipo Operación:</b></td>
						<td><select name="tipopenro" size="1" style="width:250;">
								<option value=0 selected>&nbsp;</option>
								<%Set l_rs = Server.CreateObject("ADODB.RecordSet")
								l_sql = "SELECT  * "
								l_sql  = l_sql  & " FROM buq_tipoope "
								l_sql  = l_sql  & " ORDER BY tipopedes "
								rsOpen l_rs, cn, l_sql, 0
								do until l_rs.eof		%>	
								<option value= <%= l_rs("tipopenro") %> > 
								<%= l_rs("tipopedes") %> (<%=l_rs("tipopenro")%>) </option>
								<%	l_rs.Movenext
								loop
								l_rs.Close %>
							</select>
							<script>document.datos.tipopenro.value= "<%= l_tipopenro %>"</script>
						</td>					
					</tr>					
					<tr>
						<td align="right"><b>Tipo Buque:</b></td>
						<td><select name="tipbuqnro" size="1" style="width:250;">
								<option value=0 selected>&nbsp;</option>
								<%Set l_rs = Server.CreateObject("ADODB.RecordSet")
								l_sql = "SELECT  * "
								l_sql  = l_sql  & " FROM buq_tipobuque "
								l_sql  = l_sql  & " ORDER BY tipbuqdes "
								rsOpen l_rs, cn, l_sql, 0
								do until l_rs.eof		%>	
								<option value= <%= l_rs("tipbuqnro") %> > 
								<%= l_rs("tipbuqdes") %> (<%=l_rs("tipbuqnro")%>) </option>
								<%	l_rs.Movenext
								loop
								l_rs.Close %>
							</select>
							<script>document.datos.tipbuqnro.value= "<%= l_tipbuqnro %>"</script>
						</td>					
					</tr>											
					<tr>
						<td align="right"><b>Agencia:</b></td>
						<td><select name="agenro" size="1" style="width:250;">
								<option value=0 selected>&nbsp;</option>
								<%Set l_rs = Server.CreateObject("ADODB.RecordSet")
								l_sql = "SELECT  * "
								l_sql  = l_sql  & " FROM buq_agencia "
								l_sql  = l_sql  & " ORDER BY agedes "
								rsOpen l_rs, cn, l_sql, 0
								do until l_rs.eof		%>	
								<option value= <%= l_rs("agenro") %> > 
								<%= l_rs("agedes") %> (<%=l_rs("agenro")%>) </option>
								<%	l_rs.Movenext
								loop
								l_rs.Close %>
							</select>
							<script>document.datos.agenro.value= "<%= l_agenro %>"</script>
						</td>					
					</tr>
					<tr>
					    <td align="right" nowrap width="0"><b>Comenzó;</b></td>
						<td align="left" nowrap width="0" >
						    <input type="text" name="buqfecdes" size="10" maxlength="10" value="<%= l_buqfecdes %>">
							<a href="Javascript:Ayuda_Fecha(document.datos.buqfecdes)"><img src="/serviciolocal/shared/images/cal.gif" border="0"></a>
						</td>																	
					</tr>						
					<tr>
					    <td align="right" nowrap width="0"><b>Terminó;</b></td>
						<td align="left" nowrap width="0" >
						    <input type="text" name="buqfechas" size="10" maxlength="10" value="<%= l_buqfechas %>">
							<a href="Javascript:Ayuda_Fecha(document.datos.buqfechas)"><img src="/serviciolocal/shared/images/cal.gif" border="0"></a>
						</td>																	
					</tr>
					<tr>
					    <td align="right"><b>Total Toneladas:</b></td>
						<td>
							<input type="text" name="buqton" size="10" maxlength="10" value="<%= l_buqton %>">
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

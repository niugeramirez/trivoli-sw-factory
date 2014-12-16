<% Option Explicit %>
<!--#include virtual="/serviciolocal/ess/shared/inc/const.inc"-->
<!--#include virtual="/serviciolocal/shared/db/conn_db.inc"-->
<!--#include virtual="/serviciolocal/ess/shared/inc/accesoESS.inc"-->

<!--
Archivo: ag_estudios_informales_cap_02.asp
Descripción: Abm de Estudios Informales
Autor : Lisandro Moro
Fecha: 29/01/2004
-->
<% 
on error goto 0

'Datos del formulario
Dim l_ternro
Dim l_estinffecha
Dim l_modnro
Dim l_empleg
Dim l_tipcurnro
Dim l_instnro
Dim l_estinfnro
Dim l_estinfdesabr
Dim l_estinfdesext


'ADO
Dim l_tipo
Dim l_sql
Dim l_rs
Dim l_rs1

l_tipo   = request.querystring("tipo")
l_ternro = request.querystring("ternro")

%>
<html>
<head>
<link href="../<%= c_estilo %>" rel="StyleSheet" type="text/css">
<meta http-equiv="Content-Type" content="text/html; charset=iso-8859-1">
<title> Estudios Informales - Capacitación - RHPro &reg;</title>
</head>
<script src="/serviciolocal/shared/js/fn_ayuda.js"></script>
<script src="/serviciolocal/shared/js/fn_windows.js"></script>
<script src="/serviciolocal/shared/js/fn_fechas.js"></script>
<script>
function Validar_Formulario() {

if (document.datos.estinfdesabr.value == '') {
   	alert("Debe ingresar la Descripción Abreviada");	
	document.datos.estinfdesabr.focus();
}	
else {
	if (!(validarfecha(document.datos.estinffecha))) {
	   document.datos.estinffecha.focus();
	}
	else {
		if (document.datos.tipcurnro.value == 0) {
	    	alert("Debe ingresar un Tipo de Curso");	
			document.datos.tipcurnro.focus();
		}	
		else {
			if (document.datos.instnro.value == 0) {
		    	alert("Debe ingresar una Institución");	
				document.datos.instnro.focus();
			}	
			else {
				if (document.datos.estinfdesext.value.length > 200) {
			    	alert("La Descripción Extendida es demasiado grande.");	
					document.datos.estinfdesext.focus();
				}	
				else {
				  var d=document.datos;
				  document.valida.location = "ag_estudios_informales_cap_06.asp?tipo=<%= l_tipo%>&estinfnro="+document.datos.estinfnro.value+"&estinfdesabr="+document.datos.estinfdesabr.value+"&tipcurnro="+document.datos.tipcurnro.value;			
				}  
			}
		
		}
		
	}   

}

}

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
 var jsFecha = Nuevo_Dialogo(window, '/serviciolocal/shared/js/calendar.html', 16, 15);

 if (!(jsFecha == null))
	 txt.value = jsFecha;
}

</script>
<% 
Set l_rs = Server.CreateObject("ADODB.RecordSet")
select Case l_tipo
	Case "A":
     	l_estinfdesabr  = ""
		l_tipcurnro    = 0
		l_instnro     = 0
	    l_estinffecha = ""
		l_estinfdesext = ""
	Case "M":
		l_estinfnro = request.querystring("cabnro")
		l_sql = "SELECT estinfnro , estinfdesabr, tipcurnro, instnro, estinffecha, estinfdesext"
		l_sql = l_sql & " FROM cap_estinformal "
		l_sql  = l_sql  & " WHERE estinfnro = " & l_estinfnro
		rsOpen l_rs, cn, l_sql, 0 
		if not l_rs.eof then
			l_estinfdesabr  = l_rs("estinfdesabr")
			l_tipcurnro     = l_rs("tipcurnro")
			l_instnro       = l_rs("instnro")
			l_estinffecha   = l_rs("estinffecha")
			l_estinfdesext  = l_rs("estinfdesext")
		end if

		l_rs.Close
end select

%>
<body leftmargin="0" rightmargin="0" topmargin="0" bottommargin="0" onload="javascript:document.datos.estinfdesabr.focus()">
<form name="datos" action="ag_estudios_informales_cap_03.asp?tipo=<%= l_tipo %>&ternro=<%= l_ternro %>" method="post" >
<input type="hidden" name="estinfnro" value="<%= l_estinfnro %>">
<input type="hidden" name="falresp">

<table cellspacing="0" cellpadding="0" border="0" width="100%" height="100%">
<tr>
    <th>Datos del Estudio Informal</th>
	<th colspan="3" align="right">		  
		<!--<a class=sidebtnHLP href="Javascript:ayuda('<%'= Request.ServerVariables("SCRIPT_NAME")%>');">Ayuda</a>-->
	</th>
</tr>

<tr>
    <td align="right"><b>Desc. Abreviada:</b></td>
	<td colspan="3">
		<input type="text" name="estinfdesabr" size="60" maxlength="50" value="<%= l_estinfdesabr %>">
	</td>
</tr>

<tr>
    <td align="right" nowrap width="0"><b>Fecha:</b></td>
	<td  align="left" nowrap width="0">
	    <input  type="Text" name="estinffecha" size="10" maxlength="10" value="<%= l_estinffecha %>">
		<a href="Javascript:Ayuda_Fecha(document.datos.estinffecha)"><img src="/serviciolocal/shared/images/cal.gif" border="0"></a>	
	</td>
</tr>
<tr>
    <td align="right"><b>Tipo de Curso:</b></td>
	<td>
		<select name=tipcurnro size="1" style="width:385px">
		<% If l_tipo = "A" then %>
			<option value=0 selected><< Seleccione una opción >></option>
		<% End If %>
		<%	
			Set l_rs = Server.CreateObject("ADODB.RecordSet")
			l_sql = "SELECT tipcurnro, tipcurdesabr"
			l_sql  = l_sql  & " FROM cap_tipocurso "
			l_sql  = l_sql  & " ORDER BY tipcurdesabr"
			rsOpen l_rs, cn, l_sql, 0
			do until l_rs.eof		%>	
			<option value= <%= l_rs("tipcurnro") %> > 
			<%= l_rs("tipcurdesabr") %> (<%=l_rs("tipcurnro")%>) </option>
		<%			l_rs.Movenext
			loop
			l_rs.Close %>	
		</select>
		<script> document.datos.tipcurnro.value= "<%= l_tipcurnro %>"</script>
	</td>	
</tr>

<tr>
    <td align="right"><b>Institución:</b></td>
	<td>
		<select name=instnro size="1" style="width:385px">
		<% If l_tipo = "A" then %>
			<option value=0 selected><< Seleccione una opción >></option>
		<% End If %>
		<%	
			Set l_rs = Server.CreateObject("ADODB.RecordSet")
			l_sql = "SELECT institucion.instnro,instdes  "
			l_sql = l_sql & " FROM institucion "
                        l_sql = l_sql & " WHERE instedu = -1 "
			rsOpen l_rs, cn, l_sql, 0
			do until l_rs.eof		%>	
			<option value= <%= l_rs("instnro") %> > 
			<%= l_rs("instdes") %> (<%=l_rs("instnro")%>) </option>
		<%			l_rs.Movenext
			loop
			l_rs.Close %>	
		</select>
		<script> document.datos.instnro.value= "<%= l_instnro %>"</script>
	</td>	
</tr>

<tr>
    <td align="right"><b>Desc. Extendida:</b></td>
	<td  align="left">
	    <textarea name="estinfdesext" rows="3" cols="45" maxlength="200"><%=trim(l_estinfdesext)%></textarea>
	</td>
</tr>

<tr>
    <td  colspan="4" align="right" class="th2">
		<a class=sidebtnABM href="Javascript:Validar_Formulario()">Aceptar</a>
		<a class=sidebtnABM href="Javascript:window.close()">Cancelar</a>
	</td>
</tr>
</table>
<iframe name="valida" style="visibility=hidden;" src="blanc.asp" width="100%" height="100%"></iframe> 
</form>
<%
set l_rs = nothing
'l_Cn.Close
'set l_Cn = nothing
%>
</body>
</html>

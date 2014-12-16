<% Option Explicit %>
<!--#include virtual="/serviciolocal/ess/shared/inc/const.inc"-->
<!--#include virtual="/serviciolocal/shared/db/conn_db.inc"-->
<!--#include virtual="/serviciolocal/shared/inc/fecha.inc"-->
<!--#include virtual="/serviciolocal/ess/shared/inc/accesoESS.inc"-->
<!--
-----------------------------------------------------------------------------
Archivo        : comp_lic_emp_vacacion_gti_01.asp
Descripcion    : Complemento Licencias - Vacacion
Fecha Creacion : 25/03/2004
Autor          : Scarpa D.
Modificacion   :
  29/03/2004 - Scarpa D. - Correccion en el calculo del tope de licencias por vacaciones
  18/10/2004 - Scarpa D. - Cambio en el formato de la ventana  
-----------------------------------------------------------------------------
-->
<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.0 Transitional//EN">
<%
on error goto 0

'Complemento de Vacaciones
Dim l_vacnro

Dim l_tipo
Dim l_sql
Dim l_rs
Dim l_rs1
Dim l_thnro
Dim l_emp_licnro
Dim l_tdnro
Dim l_ternro
Dim l_empleg
Dim l_desde
Dim l_corresp
Dim l_cantidad
Dim l_canttomados

Set l_rs  = Server.CreateObject("ADODB.RecordSet")

dim leg
leg = l_ess_empleg
l_ternro = l_ess_ternro

l_empleg     = leg
l_tipo       = request.queryString("tipo")
l_tdnro      = request.queryString("tdnro")
l_emp_licnro = request.queryString("emp_licnro")
l_desde      = request.queryString("desde")
l_vacnro     = request.queryString("vacnro")

Set l_rs = Server.CreateObject("ADODB.RecordSet")

select Case l_tipo
	Case "A":
	    if l_vacnro = "" then
			l_sql = "SELECT vacnro, vacdesc, vacfecdesde, vacfechasta "  
			l_sql = l_sql & " FROM  vacacion "
			l_sql = l_sql & " ORDER BY vacfecdesde DESC "
	
			rsOpen l_rs, cn, l_sql, 0 
			
			if not l_rs.eof then
			  l_vacnro = l_rs("vacnro")
			else
			  l_vacnro = "0"
			end if
			
			l_rs.close
		end if

	Case "M", "C":  'C es consulta.
		if l_vacnro = "" then ' VACACION =========================================

			l_sql = "SELECT vacnro  "
			l_sql = l_sql & " FROM lic_vacacion "
			l_sql  = l_sql  & "WHERE lic_vacacion.emp_licnro = " & l_emp_licnro

			rsOpen l_rs, cn, l_sql, 0 

			if not l_rs.eof then
				l_vacnro = l_rs("vacnro")
			end if
			
			l_rs.close
		end if	
end select

'Busco los dias correspondientes del empleado
l_corresp  = 0

if l_vacnro <> "" AND l_vacnro <> "0" then
	l_sql = "SELECT * FROM vacdiascor "
	l_sql = l_sql & " WHERE vacnro = " & l_vacnro 
	l_sql = l_sql & "   AND ternro = " & l_ternro
	
	rsOpen l_rs, cn, l_sql, 0 
	
	if not l_rs.eof then
		l_corresp   = l_rs("vdiascorcant")
	else
	    l_corresp	= 0
	end if
	
	l_rs.close
end if

'Busco la cantidad de dias tomados de la vacaciones
l_cantidad = 0

if l_vacnro <> "" AND l_vacnro <> "0" then
	l_sql =         " SELECT * "
	l_sql = l_sql & " FROM lic_vacacion "
	l_sql = l_sql & " INNER JOIN emp_lic ON emp_lic.emp_licnro = lic_vacacion.emp_licnro "
	l_sql = l_sql & " WHERE vacnro = " & l_vacnro
    l_sql = l_sql & " AND emp_lic.licestnro= 2 " 
	l_sql = l_sql & " AND emp_lic.empleado= " & l_ternro
	
	if (l_tipo ="M") then
	   l_sql = l_sql & " AND emp_lic.emp_licnro <>" & l_emp_licnro
	end if
	
	rsOpen l_rs, cn, l_sql, 0 
	
	l_cantidad = 0
	
	do until l_rs.eof 
	   l_cantidad = l_cantidad + CInt(l_rs("elcantdias"))		
	
	   l_rs.moveNext
	loop
	
	l_rs.close
end if

l_canttomados = l_cantidad

if l_vacnro = "" then
   l_vacnro = "0"	
end if
%>
<html>
<head>
<link href="../<%= c_estilo %>" rel="StyleSheet" type="text/css">
<style>
.stytttt{
  border-left-style: solid;
  border-left-width: 1px;
  border-left-color: Black;

  border-top-style: solid;
  border-top-width: 1px;
  border-top-color: Black;

  border-bottom-style: solid;
  border-bottom-width: 1px;
  border-bottom-color: Black;

  border-right-style: solid;
  border-right-width: 1px;
  border-right-color: Black;
  
}

.styfttf{
  border-bottom-style: solid;
  border-bottom-width: 1px;
  border-bottom-color: Black;

  border-right-style: solid;
  border-right-width: 1px;
  border-right-color: Black;
}

.styfftf{
  border-bottom-style: solid;
  border-bottom-width: 1px;
  border-bottom-color: Black;
}

.styftff{
  border-right-style: solid;
  border-right-width: 1px;
  border-right-color: Black;
}

</style>	

	<title>Complemento Licencia - Vacacion </title>
</head>
<script src="/serviciolocal/shared/js/fn_windows.js"></script>
<script src="/serviciolocal/shared/js/fn_ayuda.js"></script>
<script src="/serviciolocal/shared/js/fn_ay_generica.js"></script>
<script src="/serviciolocal/shared/js/fn_hora.js"></script>
<script src="/serviciolocal/shared/js/fn_fechas.js"></script>

<script>
function ValidarDatos(){
  if (document.datos.vacnro.value == ""){
     alert('Debe seleccionar un periodo de vacaciones.');
     return 0; 
  }else{
     return 1;
  }
}

function cambio(){
  window.location = '<%= "comp_lic_emp_vacacion_gti_00.asp?tipo=" & l_tipo & "&tdnro=" & l_tdnro & "&emp_licnro=" & l_emp_licnro & "&ternro=" & l_ternro & "&empleg=" & request.querystring("empleg") %>&vacnro=' + document.datos.vacnro.value;
}

function params(){
  return ('&vacnro=' + document.datos.vacnro.value);
}

function recargar(tipo,tdnro,emplicnro,ternro,empleado,desde,hasta){
  window.location.reload();
}
</script>

<body leftmargin="0" rightmargin="0" topmargin="0" bottommargin="0" onload="javascript:parent.cambioIframe();">

<form name="datos" target="vent_oculta" action="comp_lic_emp_vacacion_gti_01.asp?tipo=<%= l_tipo %>&empleg=<%= request.querystring("empleg")%>" method="post">

<input type="Hidden" name="emp_licnro" value="">
<input type="Hidden" name="tdnro" value="">
<input type="Hidden" name="tdnroant" value="">
<input type="Hidden" name="empleado" value="">
<input type="Hidden" name="elfechadesde" value="">
<input type="Hidden" name="elfechahasta" value="">
<input type="Hidden" name="elcantdias" value="">
<input type="Hidden" name="eltipo" value="">
<input type="Hidden" name="elmaxhoras" value="">
<input type="Hidden" name="seleccion" value="">
<input type="Hidden" name="seleccion1" value="">
<input type="Hidden" name="elhoradesde1" value="">
<input type="Hidden" name="elhoradesde2" value="">
<input type="Hidden" name="elhorahasta1" value="">
<input type="Hidden" name="elhorahasta2" value="">
<input type="Hidden" name="elorden" value="">
<input type="Hidden" name="licestnro" value="">

<input type="Hidden" id="hayparams" name="hayparams" value="1">

<table width="100%" border="0" CELLPADDING="0" CELLSPACING="0" height="100%">
	<tr>
		<td colspan=3 align=left><b>Licencia por Vacaciones:</b>
		</td>
	</tr>
	<tr valign="top" height="100%">
		<td align="center" colspan="3">
		    Per&iacute;odo Vacaciones:&nbsp;<br>
			<select name="vacnro" onchange="javascript:cambio();">
			<%
			l_sql = "SELECT vacnro, vacdesc, vacfecdesde, vacfechasta "  
			l_sql = l_sql & " FROM  vacacion "
			l_sql = l_sql & " ORDER BY vacfecdesde DESC "
			rsOpen l_rs, cn, l_sql, 0 
			do until l_rs.eof
			%>
				<OPTION VALUE="<%=l_rs("vacnro")%>"><%=l_rs("vacdesc")%>-<%=l_rs("vacfecdesde")%>&nbsp;al&nbsp;<%=l_rs("vacfechasta")%></OPTION>
				<%l_rs.MoveNext
			loop
			l_rs.close
			set l_rs = nothing%>
			</select>
			<script>
			document.datos.vacnro.value = "<%= l_vacnro %>";
			</script>
		</td>
	</tr>
<tr>
  <td>
     <br>
  </td>
  <td align="center" width="95%">
    <table align="center" width="100%"  class="stytttt" cellpadding="4" cellspacing="0">
		<tr>
		<td class="styfttf"><b>Licencia&nbsp;Vacaciones&nbsp;(Anual)</b></td>
		<td class="styfftf" align="right" style="padding-right:10px;"><%= l_corresp %></td>
		</tr>
		<tr>
		<td class="styfttf"><b>Licencia&nbsp;gozada</b></td>
		<td class="styfftf" align="right" style="padding-right:10px;"><%= l_canttomados %></td>
		</tr>		
		<tr>
		<td class="styftff"><b>Licencia&nbsp;pendiente&nbsp;de&nbsp;gozar</b></td>
		<td align="right" style="padding-right:10px;"><%= l_corresp - l_canttomados %></td>
		</tr>
	</table>
  </td>
  <td>
     <br>
  </td>
</tr>
<tr>
  <td colspan="3">
    <br>
  </td>
</tr>
	
</table>

</body>
</html>

<% Option Explicit %>
<!--#include virtual="/serviciolocal/ess/shared/inc/const.inc"-->
<!--#include virtual="/serviciolocal/shared/db/conn_db.inc"-->
<!--#include virtual="/serviciolocal/shared/inc/fecha.inc"-->
<!--#include virtual="/serviciolocal/ess/shared/inc/accesoESS.inc"-->
<!--
-----------------------------------------------------------------------------
Archivo       : comp_lic_emp_carga_gti_00.asp
Descripcion   : Complemento licencias
Creacion      : 24/03/2004
Autor         : Scarpa D.
Modificacion  :
  18/10/2004 - Scarpa D. - Cambio en el formato de la ventana
-----------------------------------------------------------------------------
-->
<%
on error goto 0

Dim l_rs
Dim l_sql
Dim l_tipo
Dim l_rsl
Dim l_archivo
Dim l_thnro
Dim l_emp_licnro
Dim l_tdnro
Dim l_ternro
Dim l_empleg
Dim l_desde
Dim l_hasta
Dim l_cantdiastomados
Dim l_topelic

Dim aniofin
Dim anioini
Dim aniofin2
Dim anioini2
Dim l_cantidad

Set l_rs  = Server.CreateObject("ADODB.RecordSet")

dim leg

leg = l_ess_empleg
l_ternro = l_ess_ternro

l_empleg     = leg
l_tipo       = request.queryString("tipo")
l_tdnro      = request.queryString("tdnro")
l_emp_licnro = request.queryString("emp_licnro")
l_desde      = request.queryString("desde")
l_hasta      = request.queryString("hasta")

l_archivo = ""

if l_tdnro = "" then
   l_tdnro = "0"
end if

l_cantdiastomados = 0

if CStr(l_tdnro) = "2" then

   response.redirect "comp_lic_emp_vacacion_gti_00.asp"

else

	if l_tdnro <> "" then
		'Busco el tope para el tipo de licencia
		l_sql = "SELECT tdliman FROM tipdia "
		l_sql = l_sql & " WHERE tdnro = " & l_tdnro 
	
		l_rs.Open l_sql, cn
		if not l_rs.eof then
		   if isNull(l_rs("tdliman")) then
		      l_topelic = 0
		   else
		      l_topelic = l_rs("tdliman")
		   end if
		else
		   l_topelic = 0
		end if
		l_rs.close
	end if

    if l_desde <> "" then
	
		aniofin = "31/12/" & year(CDate(l_desde))
		anioini = "01/01/" & year(CDate(l_desde))
		
		aniofin2 = CDate(aniofin)
		anioini2 = CDate(anioini)
		
		l_sql = "SELECT emp_licnro,elfechadesde,elfechahasta, elcantdias "
		l_sql = l_sql & " FROM emp_lic "
		l_sql = l_sql & " WHERE emp_lic.empleado="& l_ternro &" and ((elfechadesde >=" & cambiafecha(anioini,"YMD",true)
		l_sql = l_sql & " and elfechahasta <= " & cambiafecha(aniofin,"YMD",true) & ") "
		l_sql = l_sql & " or (elfechadesde <  " & cambiafecha(anioini,"YMD",true)
		l_sql = l_sql & " and elfechahasta <= " & cambiafecha(aniofin,"YMD",true) 
		l_sql = l_sql & " and elfechahasta >= " & cambiafecha(anioini,"YMD",true) & ") "	
		l_sql = l_sql & " or (elfechadesde >= " & cambiafecha(anioini,"YMD",true)
		l_sql = l_sql & " and elfechahasta >  " & cambiafecha(aniofin,"YMD",true) 
		l_sql = l_sql & " and elfechadesde <= " & cambiafecha(aniofin,"YMD",true) & ") "	
		l_sql = l_sql & " or (elfechadesde <  " & cambiafecha(anioini,"YMD",true)
		l_sql = l_sql & " and elfechahasta >  " & cambiafecha(aniofin,"YMD",true) & ")) "
		l_sql = l_sql & " and tdnro = " & l_tdnro
		if (l_tipo ="M") then
		l_sql = l_sql & " and emp_licnro <>" & l_emp_licnro
		end if
		
		rsOpen l_rs, cn, l_sql, 0 
		
		l_cantidad = 0
		
		do until l_rs.eof 
		
			if (DateDiff("d",CDate(l_rs("elfechadesde")), CDate(anioini2)) <= 0) and _
			   (DateDiff("d",CDate(l_rs("elfechahasta")), CDate(aniofin2)) >= 0) then
			   
			   l_cantidad = l_cantidad + CInt(l_rs("elcantdias"))
			
			else
			   if (DateDiff("d",CDate(l_rs("elfechadesde")), CDate(anioini2)) < 0) and _
			      (DateDiff("d",CDate(l_rs("elfechahasta")), CDate(aniofin2)) >= 0) and _
			      (DateDiff("d",CDate(l_rs("elfechahasta")), CDate(anioini2)) <= 0) then
				  
			      l_cantidad = l_cantidad + DateDiff("d",CDate(anioini2),CDate(l_rs("elfechahasta"))) + 1 
				  
			   else
			      if (DateDiff("d",CDate(l_rs("elfechadesde")), CDate(anioini2)) <= 0) and _
			         (DateDiff("d",CDate(l_rs("elfechahasta")), CDate(aniofin2)) < 0)  and _
			         (DateDiff("d",CDate(l_rs("elfechadesde")), CDate(aniofin2)) >= 0) then
					 
			         l_cantidad = l_cantidad + DateDiff("d",CDate(l_rs("elfechadesde")),CDate(aniofin2)) + 1
					 
				  else
			         if (DateDiff("d",CDate(l_rs("elfechadesde")), CDate(anioini2)) > 0) and _
			            (DateDiff("d",CDate(l_rs("elfechahasta")), CDate(aniofin2)) < 0) then
						
			            l_cantidad = l_cantidad + DateDiff("d",CDate(anioini2),CDate(aniofin2)) + 1
						
					 end if
				  end if
			   end if
			end if
		
		   l_rs.moveNext
		loop
		
		l_rs.close
		
		l_cantdiastomados = l_cantidad
		
		'Si es 'licencias de turismo' me fijo si existe 'licencia de turismo de trabajo'
		if CInt(l_tdnro) = 18 then
			l_sql = "SELECT emp_licnro,elfechadesde,elfechahasta, elcantdias "
			l_sql = l_sql & " FROM emp_lic "
			l_sql = l_sql & " WHERE emp_lic.empleado="& l_ternro &" and ((elfechadesde >=" & cambiafecha(anioini,"YMD",true)
			l_sql = l_sql & " and elfechahasta <= " & cambiafecha(aniofin,"YMD",true) & ") "
			l_sql = l_sql & " or (elfechadesde <  " & cambiafecha(anioini,"YMD",true)
			l_sql = l_sql & " and elfechahasta <= " & cambiafecha(aniofin,"YMD",true) 
			l_sql = l_sql & " and elfechahasta >= " & cambiafecha(anioini,"YMD",true) & ") "	
			l_sql = l_sql & " or (elfechadesde >= " & cambiafecha(anioini,"YMD",true)
			l_sql = l_sql & " and elfechahasta >  " & cambiafecha(aniofin,"YMD",true) 
			l_sql = l_sql & " and elfechadesde <= " & cambiafecha(aniofin,"YMD",true) & ") "	
			l_sql = l_sql & " or (elfechadesde <  " & cambiafecha(anioini,"YMD",true)
			l_sql = l_sql & " and elfechahasta >  " & cambiafecha(aniofin,"YMD",true) & ")) "
			l_sql = l_sql & " and tdnro = 28 " 
			
			rsOpen l_rs, cn, l_sql, 0 
			
			'Si no tiene 'licencias de turismo de trabajo' asignadas no puede tomar 'licencias de turismo'
			if l_rs.eof then
			   l_topelic = 0
			else
			   l_topelic = 5
			end if
			
			l_rs.close
		
		end if
		
	else
	    l_cantdiastomados = 0
	end if
end if
%>
<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.0 Transitional//EN">
<html>
<head>
<link href="../<%= c_estilo %>" rel="StyleSheet" type="text/css">
	<title>Complemento Licencia</title>
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
</head>
<body leftmargin="0" rightmargin="0" topmargin="0" bottommargin="0">

<script>
function ValidarDatos(){
  return 1;
}
</script>
<form name="datos" target="vent_oculta" action="comp_lic_emp_carga_gti_01.asp?tipo=<%= l_tipo %>&clase=AltaModif&empleg=<%= request.querystring("empleg")%>" method="post">
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
<!--- <input type="Hidden" name="elfechacert" value="">
<input type="Hidden" name="elfechacertcheck" value="">
 --->
<table width="100%" border="0" CELLPADDING="0" CELLSPACING="0" align="center" height="100%">
<tr>
  <td>
     <br>
  </td>
  <td align="center" width="35%">
    <table align="center" width="50%"  class="stytttt" cellpadding="4" cellspacing="0">
		<tr>
		<td class="styfttf"><b>Tope anual:</b></td>
		<td class="styfftf"><%= l_topelic%></td>
		</tr>
		<tr>
		<td class="styfttf"><b>Licencias gozadas:</b></td>
		<td class="styfftf"><%= l_cantdiastomados%></td>
		</tr>
		<tr>
		<td class="styftff"><b>Pendiente de gozar:</b></td>
		<td><%= CInt(l_topelic) - CInt(l_cantdiastomados)%></td>
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
</form>
</body>
</html>

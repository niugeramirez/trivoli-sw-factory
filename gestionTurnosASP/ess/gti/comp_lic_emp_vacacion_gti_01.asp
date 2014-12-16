<!--#include file="comp_lic_emp_carga_gti_01.asp"-->
<!--
-----------------------------------------------------------------------------
Archivo       : comp_lic_emp_vacacion_gti_01.asp
Descripcion   : Complemento licencias
Autor         : Scarpa D.
Fecha Creacion: 25/03/2004
Modificacion  :
   18/10/2004 - Scarpa D. - Correccion al guardar los datos
-----------------------------------------------------------------------------
-->
<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.0 Transitional//EN">
<html>
<head>
<link href="/serviciolocal/shared/css/tables3.css" rel="StyleSheet" type="text/css">
	<title>Complemento Licencia</title>
</head>
<body leftmargin="0" rightmargin="0" topmargin="0" bottommargin="0">
<script src="/serviciolocal/shared/js/fn_windows.js"></script>
<script src="/serviciolocal/shared/js/fn_ayuda.js"></script>
<script src="/serviciolocal/shared/js/fn_fechas.js"></script>

<table width="100%" border="0" CELLPADDING="0" CELLSPACING="0" height="100%">
<tr>
<td>
</td>
</tr>
</table>

<%
  on error goto 0
  
  Dim m_msg_final
    m_msg_final = ""
  
  Dim l_vacnro
   
  Dim l_rs
  Dim l_cm
  Dim l_sql
  Dim l_filtro

  Dim l_sin_errores
  
'---------------------------------------------------------------------------------------------------------
' FUNCION: esFeriado - calcula si una fecha es feriado
'---------------------------------------------------------------------------------------------------------
function esFeriado(dia,pais)

  Dim l_salida
  Dim m_sql
  Dim m_rs
  
  l_salida = false
  
  Set m_rs = Server.CreateObject("ADODB.RecordSet")

  m_sql =         " SELECT * FROM feriado "
  m_sql = m_sql & " WHERE feriado.ferifecha = " & cambiafecha(dia,"","")
  
  rsOpen m_rs, cn, m_sql, 0 
  
  if not m_rs.eof then
	 m_salida = ((CInt(m_rs("tipferinro")) = 1) AND (CInt(m_rs("fericodext")) = CInt(pais) ) )
  end if
  m_rs.Close
	
  esFeriado = l_salida
end function 'esFeriado(dia)


'----------------------------------------------------------------------------------------------------
'sub controlarTurismo(tipo,tdnro,ternro,desde)

'Dim z1
'Dim z2
'Dim z3

'Dim anioini
'Dim aniofin
'Dim anioini2
'Dim aniofin2

'Dim fdesde
'Dim fhasta
'Dim factual
'Dim contador
'Dim pais

'Dim dia
'Dim es_feriado
'Dim feriados
'Dim habiles
'Dim dias_tot
'Dim m_rs
'Dim m_sql
'Dim m_cm
'    m_cm = Server.CreateObject("ADODB.Command")
'const nrolictur = 18 'Codigo Interno de lic turismo en la base
'const nroturtrab = 28 'Codigo Interno de lic turismo trabajando en la base

'if CStr(tdnro) = "2" then

 '  aniofin = desde
 '  anioini = "01/01/" & year(CDate(desde))
   
'   aniofin2 = CDate(aniofin)
 '  anioini2 = CDate(anioini)
   
   'Busco por turismo trabajado
'    Set m_rs = Server.CreateObject("ADODB.RecordSet")	
	
'	m_sql = "SELECT emp_licnro"
'	m_sql = m_sql & " FROM emp_lic "
'	m_sql = m_sql & " WHERE emp_lic.empleado="& ternro &" and ((elfechadesde >=" & cambiafecha(anioini,"YMD",true)
'	m_sql = m_sql & " and elfechahasta <= " & cambiafecha(aniofin,"YMD",true) & ") "
'	m_sql = m_sql & " or (elfechadesde <  " & cambiafecha(anioini,"YMD",true)
'	m_sql = m_sql & " and elfechahasta <= " & cambiafecha(aniofin,"YMD",true)
'	m_sql = m_sql & " and elfechahasta >= " & cambiafecha(anioini,"YMD",true) & ") "	
'	m_sql = m_sql & " or (elfechadesde >= " & cambiafecha(anioini,"YMD",true)
'	m_sql = m_sql & " and elfechahasta >  " & cambiafecha(aniofin,"YMD",true)
'	m_sql = m_sql & " and elfechadesde <= " & cambiafecha(aniofin,"YMD",true) & ") "	
'	m_sql = m_sql & " or (elfechadesde <  " & cambiafecha(anioini,"YMD",true)
'	m_sql = m_sql & " and elfechahasta >  " & cambiafecha(aniofin,"YMD",true) & ")) "
'	m_sql = m_sql & " and tdnro = " & nroturtrab 'Codigo Interno de lic turismo trabajando en la base
'	if (tipo ="M") then
'		m_sql = m_sql & " and emp_licnro <>" & l_emp_licnro
'	end if
	
'	rsOpen m_rs, cn, m_sql, 0 

'	if not m_rs.eof then
'       z1 = 5
'	else
'	   z1 = 0
'	end if

'	m_rs.close
	
	'Busco turismo
	
'	z2 = 0
	
'	m_sql = "SELECT emp_licnro,elfechadesde,elfechahasta, elcantdias "
'	m_sql = m_sql & " FROM emp_lic "
'	m_sql = m_sql & " WHERE emp_lic.empleado="& ternro &" and ((elfechadesde >=" & cambiafecha(anioini,"YMD",true)
'	m_sql = m_sql & " and elfechahasta <= " & cambiafecha(aniofin,"YMD",true) & ") "
'	m_sql = m_sql & " or (elfechadesde <  " & cambiafecha(anioini,"YMD",true)
'	m_sql = m_sql & " and elfechahasta <= " & cambiafecha(aniofin,"YMD",true) 
'	m_sql = m_sql & " and elfechahasta >= " & cambiafecha(anioini,"YMD",true) & ") "	
'	m_sql = m_sql & " or (elfechadesde >= " & cambiafecha(anioini,"YMD",true)
'	m_sql = m_sql & " and elfechahasta >  " & cambiafecha(aniofin,"YMD",true) 
'	m_sql = m_sql & " and elfechadesde <= " & cambiafecha(aniofin,"YMD",true) & ") "	
'	m_sql = m_sql & " or (elfechadesde <  " & cambiafecha(anioini,"YMD",true)
'	m_sql = m_sql & " and elfechahasta >  " & cambiafecha(aniofin,"YMD",true) & ")) "
'	m_sql = m_sql & " and tdnro = " & nrolictur 'Codigo Interno de lic turismo en la base
'	if (tipo ="M") then
'		m_sql = m_sql & " and emp_licnro <>" & l_emp_licnro
'	end if
	
'	rsOpen m_rs, cn, m_sql, 0 

'	do until m_rs.eof 

'		if (DateDiff("d",CDate(m_rs("elfechadesde")), CDate(anioini2)) <= 0) and _
'		   (DateDiff("d",CDate(m_rs("elfechahasta")), CDate(aniofin2)) >= 0) then
		   
'		   z2 = z2 + CInt(m_rs("elcantdias"))
		
'		else
'		   if (DateDiff("d",CDate(m_rs("elfechadesde")), CDate(anioini2)) < 0) and _
'		      (DateDiff("d",CDate(m_rs("elfechahasta")), CDate(aniofin2)) >= 0) and _
'		      (DateDiff("d",CDate(m_rs("elfechahasta")), CDate(anioini2)) <= 0) then
			  
'		      z2 = z2 + DateDiff("d",CDate(anioini2),CDate(m_rs("elfechahasta"))) + 1 
			  
'		   else
'		      if (DateDiff("d",CDate(m_rs("elfechadesde")), CDate(anioini2)) <= 0) and _
'		         (DateDiff("d",CDate(m_rs("elfechahasta")), CDate(aniofin2)) < 0)  and _
'		         (DateDiff("d",CDate(m_rs("elfechadesde")), CDate(aniofin2)) >= 0) then
				 
'		         z2 = z2 + DateDiff("d",CDate(m_rs("elfechadesde")),CDate(aniofin2)) + 1
				 
'			  else
'		         if (DateDiff("d",CDate(m_rs("elfechadesde")), CDate(anioini2)) > 0) and _
'		            (DateDiff("d",CDate(m_rs("elfechahasta")), CDate(aniofin2)) < 0) then
					
'		            z2 = z2 + DateDiff("d",CDate(anioini2),CDate(aniofin2)) + 1
					
'				 end if
'			  end if
'		   end if
'		end if

'	    m_rs.moveNext
'	loop

'	m_rs.close
	
'	z3 = z1 - z2
	
'	if z3 >= 0 then
'	   if (1 <= z3) and ( z3 <= 5) then
'	      set m_cm = Server.CreateObject("ADODB.Command")
'	      m_msg_final = "Se cargarán las Licencias por Turismo Pendiente: " & z3 & " dias."
		  
'          fdesde = CDate(desde)
'          fhasta = DateAdd("d", z3 - 1, CDate(desde))
		  
		  'Obtengo el pais en el que estoy
'		  m_sql =         " SELECT * FROM pais "
'		  m_sql = m_sql & " WHERE pais.paisdef = -1 " 
		
'		  rsOpen m_rs, cn, m_sql, 0 
		
'		  if not m_rs.eof then
'		    pais = m_rs("paisnro") 
'		  else
'		    pais = 1
'		  end if

'		  m_rs.close

		  'Busco la cantidad de dias habiles y feriados
'		  contador = DateDiff("d",fdesde,fhasta)
'		  factual = CDate(fdesde)
'		  do 
'			 dias_tot = dias_tot + 1	 
'		   	 if esFeriado(factual,pais) then	 
			    'es un feriado 
'		        feriados = feriados + 1	   
'			 else
			    'Dependiendo si es dia laborable o no
'		        if (WeekDay(factual) = 1) OR (WeekDay(factual) = 7) then
'		           feriados = feriados + 1	   		
'				else
'				   habiles = habiles + 1
'				end if
'			 end if
'		 	 factual = DateAdd("d", 1, factual)   
'			 contador = contador - 1
'		  loop while contador >= 0		  
		  
		  'Inserto los datos en la BD
'		  m_sql = "insert into emp_lic "
'		  m_sql = m_sql & "(emp_licnro, tdnro, empleado, elfechadesde, elfechahasta, elcantdias, elmaxhoras, elorden, eltipo, licestnro, eldiacompleto, elcanthrs, elcantdiasfer, elcantdiashab ) "
'		  m_sql = m_sql & "values (" & nrolictur & ", " & ternro & ", " & cambiafecha(fdesde,"YMD",true) & ", " & cambiafecha(fhasta,"YMD",true)
'		  m_sql = m_sql & ", " & dias_tot & ", " & l_elmaxhoras & ", " & l_elorden & ", 1 , 1, -1, 0, " & feriados & ", " & habiles & ")"
		  
'		  m_cm.activeconnection = Cn
'		  m_cm.CommandText = m_sql
'		  cmExecute m_cm, m_sql, 0

'	   end if
'	end if

'end if

'end sub 'controlarTurismo(tdnro,ternro)




  l_vacnro = Request.Form("vacnro")

  l_sin_errores = true
  
  obtDatosFormulario()
  
  if controlarLicencias(request.queryString("tipo")) then
     Response.write "<script>alert('Esta Licencia se superpone con otras cargadas anteriormente.');history.back();</script>"
    l_sin_errores = false
  end if
  
  if controlarRelaciones() then
	Response.write "<script>alert('No puede cambiar el tipo de esta Licencia. Tiene Licencias vinculadas.');history.back();</script>"
    l_sin_errores = false	
  end if
  
  if l_sin_errores then
	  cn.beginTrans
		  set l_cm = Server.CreateObject("ADODB.Command")
	      Set l_rs = Server.CreateObject("ADODB.RecordSet")
		  
	      'Genero la Licencia
	      l_sql = generarSQLLic(request.queryString("tipo"))
		  l_cm.activeconnection = Cn
		  l_cm.CommandText = l_sql
		  cmExecute l_cm, l_sql, 0

		  if request.queryString("tipo") = "A" then
	  		 ' Busco el nro de la licencia
		      l_sql = fsql_seqvalue("codigo","emp_lic")
   	          rsOpen l_rs, cn, l_sql, 0
	          l_emp_licnro=l_rs("codigo")
              l_rs.Close
		  end if

 		  'Genero la sql del complemento de acuerdo al caso
   		  if request.queryString("tipo") = "A" then
				l_sql = "INSERT INTO lic_vacacion "
				l_sql = l_sql & "(emp_licnro, vacnro, licvacmanual) "
				l_sql = l_sql & "values (" & l_emp_licnro &", " 
				if l_vacnro = "" then
				    l_sql = l_sql & "0,-1)"
				else
				    l_sql = l_sql & l_vacnro & ", -1)"
				end if
		  else
		     if l_tdnro <> l_tdnroant then
			 
			    'Borro el complemento anterior
				borrarComplementos(l_tdnroant)

				l_sql = "INSERT INTO lic_vacacion "
				l_sql = l_sql & "(emp_licnro, vacnro, licvacmanual) "
				l_sql = l_sql & "values (" & l_emp_licnro &", " 
				if l_vacnro = "" then
				    l_sql = l_sql & "0,-1)"
				else
				    l_sql = l_sql & l_vacnro & ", -1)"
				end if
			 else
				if l_vacnro = "" then
				   l_vacnro = "0"
				end if

				l_sql = "UPDATE lic_vacacion SET "
				l_sql = l_sql & " vacnro	 = " & l_vacnro & ", "  
				l_sql = l_sql & " licvacmanual	 = -1  "  
				l_sql = l_sql & " WHERE emp_licnro = " & l_emp_licnro
			 end if
		  end if

		  'Guardo los datos en la BD 
		  l_cm.activeconnection = Cn
		  l_cm.CommandText = l_sql
		  cmExecute l_cm, l_sql, 0

		  'Genero la Justificacion
	      l_sql = generarSQLJust(request.queryString("tipo"))
		  l_cm.activeconnection = Cn
		  l_cm.CommandText = l_sql
		  cmExecute l_cm, l_sql, 0
		  
		  'call controlarTurismo(request.queryString("tipo"),l_tdnro,l_empleado,l_elfechadesde)
		  
		  'Genero las Firmas1 
		  if (l_cysfirmas1 <> "") then
		      l_sql = generarSQLFirmas1()
			  l_cm.activeconnection = Cn
			  l_cm.CommandText = l_sql
		      cmExecute l_cm, l_sql, 0
		  end if

		  'Genero las Firmas
		  if (l_cysfirmas <> "") then
		      l_sql = generarSQLFirmas()
			  l_cm.activeconnection = Cn
			  l_cm.CommandText = l_sql
		      cmExecute l_cm, l_sql, 0
		  end if
		  
	  cn.CommitTrans
	
	  Set cn = Nothing
	  Set l_cm = Nothing
%>	  

<script>
  <%if m_msg_final <> "" then%>
  alert('<%= m_msg_final%>');
  <%end if%>
   //abrirVentanaH('postmail.asp',100,100);
   alert('Operación Realizada.');

  window.opener.opener.ifrm.location.reload();
  window.opener.close();    
  window.close();
</script>
<%
  end if

%>

</body>
</html>

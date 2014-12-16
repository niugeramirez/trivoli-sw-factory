<% Option Explicit %>
<!--#include virtual="/serviciolocal/shared/inc/sec.inc"-->
<!--#include virtual="/serviciolocal/shared/inc/const.inc"-->
<!--#include virtual="/serviciolocal/shared/db/conn_db.inc"-->
<!--#include virtual="/serviciolocal/shared/inc/fecha.inc"-->
<!--#include virtual="/serviciolocal/shared/inc/sqls.inc"-->
<!--#include virtual="/serviciolocal/shared/inc/liccantidaddias.inc"-->
<!--#include virtual="/serviciolocal/shared/inc/adovbs.inc"-->
<!--#include virtual="/serviciolocal/ess/shared/inc/accesoESS.inc"-->
<!--
-----------------------------------------------------------------------------
Archivo       : comp_lic_emp_carga_gti_01.asp
Descripcion   : Complemento licencias
Creacion      : 24/03/2004
Autor         : Scarpa D.
Modificacion  :
  06/05/2004 - Scarpa D. - Se quitarin los campos de licencias parciales
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
const l_valornulo = "null"

'Datos 
Dim l_emp_licnro
Dim l_tdnro
Dim l_tdnroant
Dim l_empleado
Dim l_elfechadesde
Dim l_elfechahasta
Dim l_elcantdias
Dim l_eltipo
Dim l_elhoradesde
Dim l_elhorahasta
Dim l_elorden
Dim l_elmaxhoras
Dim l_licestnro
Dim l_Cysfirmas
Dim l_Cysfirmas1
'Dim l_elfechacert
'Dim l_elfechacertcheck

'Variables locales
Dim l_clase
Dim m_filtro
Dim m_sql
Dim m_rs
Dim m_cm

l_clase = request.queryString("clase")

'----------------------------------------------------------------------------------------------------
'Borra los complementos de acuerdo a un tdnro
sub borrarComplementos(tdnro)

    Set m_cm = Server.CreateObject("ADODB.Command")

	if tdnro = 2 then

		m_sql = "DELETE FROM lic_vacacion "
		m_sql = m_sql & " WHERE emp_licnro = " & l_emp_licnro
		
		m_cm.activeconnection = Cn
		m_cm.CommandText = m_sql
		cmExecute m_cm, m_sql, 0

	end if 

end sub 'borrarComplementosLocal(tdnro)


'----------------------------------------------------------------------------------------------------
'Obtiene los datos del formulario
sub obtDatosFormulario()

l_emp_licnro   = request.Form("emp_licnro")
l_tdnro		   = request.Form("tdnro")
l_tdnroant     = request.Form("tdnroant")
'l_empleado     = request.Form("empleado")
l_empleado     = l_ess_ternro
l_elfechadesde = request.Form("elfechadesde")
l_elfechahasta = request.Form("elfechahasta")
l_elcantdias   = request.Form("elcantdias")
l_eltipo       = request.Form("eltipo")
l_elmaxhoras   = request.Form("elmaxhoras")
l_licestnro    = request.Form("licestnro")
l_cysfirmas    = request.Form("seleccion")
l_cysfirmas1   = request.Form("seleccion1")
'l_elfechacert  = request.Form("elfechacert")
'l_elfechacertcheck = request.Form("elfechacertcheck")

if trim(l_licestnro) = "" then
   l_licestnro = "1"
end if

'if CBool(l_elfechacertcheck) then
'	if trim(l_elfechacert) = "" then
'       l_elfechacert = "null"	
'	else
'       l_elfechacert = cambiafecha(l_elfechacert,"YMD",true) 	   
'	end if
'else
'   l_elfechacert = "null"
'end if

'Custom ABN
l_eltipo = "1"
l_elmaxhoras=""

if l_elmaxhoras="" then
	l_elmaxhoras= "0"
end if 

if l_eltipo = "1" then
	l_elhoradesde = l_valornulo
	l_elhorahasta = l_valornulo
	l_elorden = l_valornulo
end if
if l_eltipo = "2" then 
	l_elhoradesde = "'"&request.Form("elhoradesde1") & request.Form("elhoradesde2")&"'"
	l_elhorahasta = "'"&request.Form("elhorahasta1") & request.Form("elhorahasta2")&"'"
	l_elorden = l_valornulo
end if 
if l_eltipo = "3" then 
	l_elhoradesde = l_valornulo
	l_elhorahasta = l_valornulo
	l_elorden = request.Form("elorden")
end if 

end sub 'obtDatosFormulario()

'----------------------------------------------------------------------------------------------------
'Controla que no exista otra licencia en el mismo periodo de fecha
function controlarLicencias(tipo)
    Dim l_salida

	Set m_rs = Server.CreateObject("ADODB.RecordSet")	
	
	m_sql = "SELECT emp_licnro, eltipo, elhoradesde, elhorahasta "
	m_sql = m_sql & " FROM emp_lic "
	m_sql = m_sql & " WHERE emp_lic.empleado="& l_empleado &" and ((elfechadesde >=" & cambiafecha(l_elfechadesde,"YMD",true)
	m_sql = m_sql & " and elfechadesde <=" & cambiafecha(l_elfechahasta,"YMD",true) & ") "
	m_sql = m_sql & " or (elfechahasta >=" & cambiafecha(l_elfechadesde,"YMD",true)
	m_sql = m_sql & " and elfechahasta <=" & cambiafecha(l_elfechahasta,"YMD",true) & ") "
	m_sql = m_sql & " or (elfechadesde <=" & cambiafecha(l_elfechadesde,"YMD",true)
	m_sql = m_sql & " and elfechahasta >=" & cambiafecha(l_elfechadesde,"YMD",true) & ") "
	m_sql = m_sql & " or (elfechadesde <=" & cambiafecha(l_elfechahasta,"YMD",true)
	m_sql = m_sql & " and elfechahasta >=" & cambiafecha(l_elfechahasta,"YMD",true) & ")) "
	if (tipo ="M") then
		m_sql = m_sql & " and emp_licnro <>" & l_emp_licnro
	end if
	if l_eltipo="2" then
		m_sql = m_sql & " and (eltipo =1 " 
		m_sql = m_sql & " or ( eltipo=2 and ( (elhoradesde <=" & l_elhorahasta & " and elhorahasta>="& l_elhorahasta & " ) "
		m_sql = m_sql & " or (elhoradesde <=" & l_elhoradesde & " and elhorahasta>="& l_elhoradesde & " ) "
		m_sql = m_sql & " or (elhoradesde >=" & l_elhoradesde & " and elhoradesde<="& l_elhorahasta & " )))) "
	end if
	if l_eltipo="3" then
		m_sql = m_sql & " and eltipo =1 " 
	end if
	
	rsOpen m_rs, cn, m_sql, 0 
	l_salida = false
	if not m_rs.eof then
       l_salida = true
	end if
	m_rs.close
	
	controlarLicencias = l_salida

end function 'controlarLicencias()

'----------------------------------------------------------------------------------------------------
'Controla las relacion con otra licencias
function controlarRelaciones()
    Dim l_salida
	
	l_salida = false

	if (l_tdnroant <> "") then
		if (l_tdnro<>l_tdnroant) and (l_tdnroant=9 or l_tdnroant=13 or l_tdnroant=14) then
			Set m_rs = Server.CreateObject("ADODB.RecordSet")
			m_sql = "SELECT * FROM emp_lic "
			m_sql = m_sql & " where emp_licnro = " & l_emp_licnro
			rsOpen m_rs, cn, m_sql, 0 
			if not m_rs.eof then
				if m_rs("licnrosig") <> "" then
					l_salida = true
				end if	
			end if
			m_rs.close
		end if
	end if
	
	controlarRelaciones = l_salida

end function 'controlarRelaciones()

'----------------------------------------------------------------------------------------------------

'Genera la sql de la licencia de acuerdo si es un alta o una modificacion
function generarSQLLic(tipo)

if tipo = "A" then 
	m_sql = "insert into emp_lic "
	'm_sql = m_sql & "(emp_licnro, tdnro, empleado, elfechadesde, elfechahasta, elcantdias, elmaxhoras, elorden, eltipo, elhoradesde, elhorahasta, elfechacert, licestnro ) "
	m_sql = m_sql & "(tdnro, empleado, elfechadesde, elfechahasta, elcantdias, elmaxhoras, elorden, eltipo, elhoradesde, elhorahasta, licestnro ) "
	m_sql = m_sql & "values (" & l_tdnro &", " & l_empleado & ", " & cambiafecha(l_elfechadesde,"YMD",true) & ", " & cambiafecha (l_elfechahasta,"YMD",true)
	'm_sql = m_sql & ", " & l_elcantdias & ", " & l_elmaxhoras & ", " & l_elorden & ", " & l_eltipo & ", " & l_elhoradesde & ", " & l_elhorahasta & ", " & l_elfechacert & ", 1)"
	m_sql = m_sql & ", " & l_elcantdias & ", " & l_elmaxhoras & ", " & l_elorden & ", " & l_eltipo & ", " & l_elhoradesde & ", " & l_elhorahasta & ", " & l_licestnro & ")"
else
	m_sql = "update emp_lic "
	m_sql = m_sql & "set  tdnro="& l_tdnro & ", empleado=" & l_empleado & ", elfechadesde="& cambiafecha(l_elfechadesde,"YMD",true) & ", elfechahasta =" & cambiafecha(l_elfechahasta,"YMD",true) & ", elcantdias =" & l_elcantdias 
	'm_sql = m_sql & ", elmaxhoras = " & l_elmaxhoras & ", elorden =" & l_elorden & ", eltipo =" & l_eltipo & ", elhoradesde = " & l_elhoradesde & ", elhorahasta=" & l_elhorahasta & ", elfechacert =" & l_elfechacert
	m_sql = m_sql & ", elmaxhoras = " & l_elmaxhoras & ", elorden =" & l_elorden & ", eltipo =" & l_eltipo & ", elhoradesde = " & l_elhoradesde & ", elhorahasta=" & l_elhorahasta & ", licestnro=" & l_licestnro
	m_sql = m_sql & " where emp_licnro = " & l_emp_licnro
end if

    generarSQLLic = m_sql
	
end function 'generarSQLLic(tipo)

'----------------------------------------------------------------------------------------------------
'Genera la sql para las justificaciones
function generarSQLJust(tipo)

	if tipo = "A" then 
	    m_sql = "INSERT INTO gti_justificacion (jusanterior,juscodext,jusdesde,jusdiacompleto,jushasta,jussigla,jussistema,ternro,tjusnro,turnro,jushoradesde,jushorahasta,juseltipo,juselorden,juselmaxhoras )" & _
	                    " VALUES(1," & l_emp_licnro & "," & cambiafecha(l_elfechadesde,"YMD",true) & ",-1," & cambiafecha(l_elfechahasta,"YMD",true) & ",'LIC',-1," & l_empleado & ",1,0," & l_elhoradesde & "," & l_elhorahasta & "," & l_eltipo & "," & l_elorden & "," & l_elmaxhoras & ")"
	
	else
		m_sql = "UPDATE gti_justificacion SET jusanterior = 1, jusdesde = " & cambiafecha(l_elfechadesde,"YMD",true) & ", jusdiacompleto = -1, jushasta = " & cambiafecha(l_elfechahasta,"YMD",true) & ",jussistema = -1"  & _
	                      ", tjusnro = 1 ,turnro = 0 ,jushoradesde = " & l_elhoradesde & ", jushorahasta = " & l_elhorahasta & ", juseltipo = " & l_eltipo & ", juselorden = " & l_elorden & ", juselmaxhoras = " & l_elmaxhoras	&_
						 " WHERE ternro= " & l_empleado & " and jussigla = 'LIC' and juscodext = " & l_emp_licnro
	end if

    generarSQLJust = m_sql

end function 'generarSQLJust(tipo)

'----------------------------------------------------------------------------------------------------
'Genera la SQL de las firmas1
function generarSQLFirmas1()

	if (l_cysfirmas1 <> "") then
	  if inStr(l_cysfirmas1,"@@@") <> 0 then
	    l_cysfirmas1 = left(l_cysfirmas1,inStr(l_cysfirmas1,"@@@")-1) & l_emp_licnro & mid(l_cysfirmas1,inStr(l_cysfirmas1,"@@@")+3)
	  end if
	end if  
	
	generarSQLFirmas1 = l_cysfirmas1
	
end function 'generarSQLFirmas1()
	

'----------------------------------------------------------------------------------------------------
'Genera la SQL de las firmas
function generarSQLFirmas()
	
	if (l_cysfirmas <> "") then
	  if inStr(l_cysfirmas,"@@@") <> 0 then
	    l_cysfirmas = left(l_cysfirmas,inStr(l_cysfirmas,"@@@")-1) & l_emp_licnro & mid(l_cysfirmas,inStr(l_cysfirmas,"@@@")+3)
	  end if
	end if  
	
	generarSQLFirmas = l_cysfirmas 

end function 'generarSQLFirmas()

'----------------------------------------------------------------------------------------------------
'Guarda los datos si es un alta o una modificacion
'----------------------------------------------------------------------------------------------------
if l_clase = "AltaModif" then
  Dim m_sin_errores

  m_sin_errores = true
  
  obtDatosFormulario()
  
  if controlarLicencias(request.queryString("tipo")) then
    Response.write "<script>alert('Esta Licencia se superpone con otras cargadas anteriormente.');history.back();</script>"
    m_sin_errores = false
  end if
  
  if controlarRelaciones() then
	Response.write "<script>alert('No puede cambiar el tipo de esta Licencia. Tiene Licencias vinculadas.');history.back();</script>"
    m_sin_errores = false	
  end if
  
  if m_sin_errores then
	  cn.beginTrans
		  set m_cm = Server.CreateObject("ADODB.Command")
		  
		  'Borro los complementos de la licencia anterior
   		  if request.queryString("tipo") = "M" then		  
		     borrarComplementos(l_tdnroant)
		  end if
		  
	      'Genero la Licencia
	      m_sql = generarSQLLic(request.queryString("tipo"))
		  m_cm.activeconnection = Cn
		  m_cm.CommandText = m_sql
		  cmExecute m_cm, m_sql, 0
		  
		  if request.queryString("tipo") = "A" then
		      m_sql = fsql_seqvalue("codigo","emp_lic")
   	          rsOpen m_rs, cn, m_sql, 0
	          l_emp_licnro=m_rs("codigo")
              m_rs.Close
		  end if
		  
		  'Genero la Justificacion
	      m_sql = generarSQLJust(request.queryString("tipo"))
		  m_cm.activeconnection = Cn
		  m_cm.CommandText = m_sql
		  cmExecute m_cm, m_sql, 0
		  
		  'Genero las Firmas1 
		  if (l_cysfirmas1 <> "") then
		      m_sql = generarSQLFirmas1()
			  m_cm.activeconnection = Cn
			  m_cm.CommandText = m_sql
		      cmExecute m_cm, m_sql, 0
		  end if
		  
		  'Genero las Firmas
		  if (l_cysfirmas <> "") then
		      m_sql = generarSQLFirmas()
			  m_cm.activeconnection = Cn
			  m_cm.CommandText = m_sql
		      cmExecute m_cm, m_sql, 0
		  end if
		  
	  cn.CommitTrans

	  Set cn = Nothing
	  Set m_cm = Nothing

%>

<script>
 // abrirVentanaH('postmail.asp',100,100);	
  alert('Operación Realizada.');
  window.opener.opener.ifrm.location.reload();
  window.opener.close();    
  window.close();
</script>
<%	

  end if

end if

%>

</body>
</html>

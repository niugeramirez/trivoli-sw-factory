<%  Option Explicit %>
<!--#include virtual="/serviciolocal/shared/inc/sec.inc"-->
<!--#include virtual="/serviciolocal/shared/inc/const.inc"-->
<!--#include virtual="/serviciolocal/shared/db/conn_db.inc"-->
<!--#include virtual="/serviciolocal/shared/inc/fecha.inc"-->
<!--#include virtual="/serviciolocal/shared/inc/util.inc"-->
<!--#include virtual="/serviciolocal/shared/inc/sqls.inc"-->
<!--#include virtual="/serviciolocal/shared/inc/tope_vales_liq.inc"-->
<!--
-----------------------------------------------------------------------------
Archivo        : vales_liq_03.asp
Descripcion    : Modulo que se encarga de las Altas/Modif de vales
Creador        : Scarpa D.
Fecha Creacion : 12/01/2004
Modificación   :
	09-02-04 - F.Favre - Cierra la ventana llamadora, solamente si es una modificacion. 
	25-06-04 - F.Favre - El campo valmonto no se guardaba correctamente. 
  Modificado  : 12/09/2006 Raul Chinestra - se agregó Vales en Autogestión   
  				23/02/2007 - Martin Ferraro - Guardar el apellido y nombre de quien da de ALTA los vales
-----------------------------------------------------------------------------
-->
<% 
'on error goto 0

' Declaracion de Variables locales  -------------------------------------
Dim l_tipo
Dim l_cm
Dim l_sql
dim l_rs

Dim l_cysfirmas
Dim l_cysfirmas1

Dim l_error
Dim l_emp_err
    l_error = false
    l_emp_err = ""

Dim l_valnro
Dim l_empleado
Dim l_ppagnro
Dim l_monnro
Dim l_valmonto
Dim l_valfecped
Dim l_valfecprev
Dim l_pliqnro
Dim l_valdesc
Dim l_pliqdto
Dim l_pronro
Dim l_tvalenro
Dim l_valrevis

Dim l_valusuario

Set l_rs = Server.CreateObject("ADODB.RecordSet")

' traer valores del form de alta/modificacion -------------------------------------

l_tipo = request("tipo")

l_valnro     = request.Form("valnro")
l_empleado   = request.Form("ternro")
l_monnro     = request.Form("monnro")
l_valmonto   = request.Form("valmonto2")
l_valfecped  = getFecha("valfecped")
l_valfecprev = getFecha("valfecprev")
l_pliqnro    = request.Form("pliqnro")
l_valdesc    = getString("valdesc")
l_pliqdto    = request.Form("pliqdto")
l_tvalenro   = request.Form("tvalenro2")
l_valrevis   = getCheckbox("valrevis")

set l_cm = Server.CreateObject("ADODB.Command")
Set l_rs = Server.CreateObject("ADODB.RecordSet")

cn.beginTrans

'Inicializo la busqueda del tope
inicializarObtTope l_pliqdto

l_valusuario = ""

if m_tope_restrictivo then
  
  if CDbl(l_valmonto) > CDbl(obtenerTope(l_empleado,l_pliqdto,l_valnro)) then
     l_error = true
	 if l_emp_err = "" then
        l_emp_err = l_empleado
	 else
        l_emp_err = l_emp_err & "," & l_empleado
	 end if
  end if
end if

if not l_error then



		if l_tipo = "A" then 
		
			'23/02/2007 - Busco los datos del usuario registrado
			l_sql = " SELECT terape, terape2, ternom, ternom2"
			l_sql = l_sql & " FROM empleado"
			l_sql = l_sql & " WHERE empleado.empleg = " & Session("empleg")
			rsOpen l_rs, cn, l_sql, 0 
			if not l_rs.eof then
				l_valusuario = l_rs("terape") & " " & l_rs("ternom")
			end if
			l_rs.close
			
			l_sql = "INSERT INTO vales "
			l_sql = l_sql & "(empleado,ppagnro,monnro,valmonto, "
			l_sql = l_sql & " valfecped,valfecprev,pliqnro,valdesc, "
			l_sql = l_sql & " pliqdto,pronro,tvalenro,valrevis,valusuario) "
			l_sql = l_sql & " VALUES (" 
			l_sql = l_sql & l_empleado   & ","
			l_sql = l_sql & "NULL"       & ","
			l_sql = l_sql & l_monnro     & ","
			l_sql = l_sql & l_valmonto   & ","
			l_sql = l_sql & l_valfecped  & ","	
			l_sql = l_sql & l_valfecprev & ","	
			l_sql = l_sql & l_pliqnro    & ","	
			l_sql = l_sql & l_valdesc    & ","	
			l_sql = l_sql & l_pliqdto    & ","	
			l_sql = l_sql & "NULL"       & ","	
			l_sql = l_sql & l_tvalenro   & ","	
			l_sql = l_sql & 0   & ","
			l_sql = l_sql & "'" & l_valusuario & "'"	
			l_sql = l_sql & ")"	
			
			l_cm.activeconnection = Cn
			l_cm.CommandText = l_sql
			cmExecute l_cm, l_sql, 0	
		
		else
		
			l_sql = "UPDATE vales SET "
			l_sql = l_sql & " empleado="   & l_empleado   & ","
			l_sql = l_sql & " monnro="     & l_monnro     & ","
			l_sql = l_sql & " valmonto="   & l_valmonto   & ","
			l_sql = l_sql & " valfecped="  & l_valfecped  & ","	
			l_sql = l_sql & " valfecprev=" & l_valfecprev & ","	
			l_sql = l_sql & " pliqnro="    & l_pliqnro    & ","	
			l_sql = l_sql & " valdesc="    & l_valdesc    & ","	
			l_sql = l_sql & " pliqdto="    & l_pliqdto    & ","	
			l_sql = l_sql & " tvalenro="   & l_tvalenro
			l_sql = l_sql & " WHERE valnro = "  & l_valnro
			
			l_cm.activeconnection = Cn
			l_cm.CommandText = l_sql
			'response.write l_sql
			cmExecute l_cm, l_sql, 0	
		
		end if
end if

cn.CommitTrans

set l_rs = Nothing

if l_error then
   	Response.write "<script>alert('El monto del vale supera el tope establecido.');window.close();</script>"
else
   	Response.write "<script>alert('Operación Realizada.');window.opener.opener.ifrm.location.reload();</script>"
    Response.write "<script>window.opener.close();</script>"
	Response.write "<script>window.close();</script>"
end if
%>

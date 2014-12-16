<% Option Explicit %>
<!--#include virtual="/serviciolocal/shared/db/conn_db.inc"-->
<!--#include virtual="/serviciolocal/shared/inc/const.inc"-->
<!--#include virtual="/serviciolocal/shared/inc/adovbs.inc"-->
<!--#include virtual="/serviciolocal/ess/shared/inc/accesoESS.inc"-->
<!--
-----------------------------------------------------------------------------
Archivo       : licencias_emp_gti_04.asp
Descripcion   : Modulo que se encarga de la baja de licencias
Creacion      : 24/03/2004
Autor         : Scarpa D.
Modificacion  :
-----------------------------------------------------------------------------
-->
<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.0 Transitional//EN">
<html>
<head>
<link href="/serviciolocal/shared/css/tables3.css" rel="StyleSheet" type="text/css">
	<title>Baja Licencia</title>
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
 
 Dim l_cm
 Dim l_sql
 Dim l_emp_licnro
 Dim l_empleado
 Dim l_filtro
 Dim l_rs
 Dim l_tdnro
 Dim l_noborrar
 
 Dim l_desde
 Dim l_hasta
 
 l_emp_licnro = request.querystring("cabnro")
 l_filtro = request.querystring("filtro")
 
'tomar el tipo de licencia para COMPLEMENTOS
 Set l_rs = Server.CreateObject("ADODB.RecordSet") 
 l_rs.CursorType = adOpenKeyset  
 l_sql = "SELECT emp_licnro, tdnro, empleado, elfechadesde, elfechahasta "
 l_sql = l_sql & " FROM emp_lic "
 l_sql  = l_sql  & "WHERE emp_licnro = " & l_emp_licnro
 rsOpen l_rs, cn, l_sql, 0 
 if not l_rs.eof then
	l_empleado = l_rs("empleado")
	l_tdnro = l_rs("tdnro")
'	if isNull(l_rs("licnrosig")) then
	  l_noborrar = false
'	else
'	  if (l_tdnro=9 or l_tdnro=14 or l_tdnro=13) and (( CStr(l_rs("licnrosig")) <> "" ) AND ( CStr(l_rs("licnrosig")) <>"0" ) ) then
'		 l_noborrar = true
'	  end if
'	end if
	'Guardo el rango de fechas para el procesamiento online 
	l_desde = l_rs("elfechadesde")
	l_hasta = l_rs("elfechahasta")
 end if
 l_rs.close
 
if l_noborrar then
	response.write "<script>alert('Esta licencia tiene licencias vinculadas. No puede borrarla.');window.close();</script>"
	response.end
 else
	
' Se verifica que la licencia no tenga Pagos/Dto asociados
'	l_noborrar = false
'	l_sql = "SELECT * "
'	l_sql = l_sql & " FROM vacpagdesc "
'	l_sql  = l_sql  & "WHERE emp_licnro = " & l_emp_licnro
'	rsOpen l_rs, cn, l_sql, 0 
'	if not l_rs.eof then
'		l_noborrar = true
'	end if
'	l_rs.close
	
	if l_noborrar then
		response.write "<script>alert('Esta licencia tiene Pagos/Dto asociados. No puede borrarla.');window.close();</script>"
		response.end
	else
		
		Dim l_tipAutorizacion  'Es el tipo del circuito de firmas
		Dim l_HayAutorizacion  'Es para ver si las autorizaciones estan activas
		Dim l_PuedeVer         'Es para ver si las autorizaciones estan activas
		
		l_tipAutorizacion = 6  'Es del tipo licencias
		
		'Set l_rs = Server.CreateObject("ADODB.RecordSet")
		l_sql = "select * from cystipo "
		l_sql = l_sql & "where (cystipo.cystipact = 1) and cystipo.cystipnro = " & l_tipAutorizacion 
		rsOpen l_rs, cn, l_sql, 0 
		
		l_HayAutorizacion = not l_rs.eof
		l_rs.close
		
		if l_HayAutorizacion then
			
			l_sql = "select cysfirautoriza, cysfirsecuencia, cysfirdestino from cysfirmas "
		  	l_sql = l_sql & "where cysfirmas.cystipnro = " & l_tipAutorizacion & " and cysfirmas.cysfircodext = '" & l_emp_licnro & "' " 
		  	l_sql = l_sql & "order by cysfirsecuencia desc"
			
		  	rsOpen l_rs, cn, l_sql, 0 
			
		  	l_PuedeVer = False
		  	' borrar =====================================================
		  	l_PuedeVer = true 
			
		  	if not l_rs.eof then
		    	if (l_rs("cysfirautoriza") = session("UserName")) or (l_rs("cysfirdestino") = session("UserName")) then 
			   	'Es una modificación del ultimo o es el nuevo que autoriza 
		       		l_PuedeVer = True 
		   		end if
		  	end if
		  	l_rs.close
		  	If not l_PuedeVer then
		    	response.write "<script>alert('No esta autorizado a ver o modificar este registro.');window.close()</script>"
				response.end
		  	End if
		End if
		
		
		cn.beginTrans
		set l_cm = Server.CreateObject("ADODB.Command")
		
		if l_tdnro = 11 then ' MATERNIDAD
			l_sql = "DELETE FROM lic_mater " 
			l_sql = l_sql & "WHERE emp_licnro = " & l_emp_licnro
			l_cm.activeconnection = Cn
			l_cm.CommandText = l_sql
			cmExecute l_cm, l_sql, 0
		end if
		if l_tdnro = 12 then ' MATERNIDAD
			l_sql = "DELETE FROM lic_lact " 
			l_sql = l_sql & "WHERE emp_licnro = " & l_emp_licnro
			l_cm.activeconnection = Cn
			l_cm.CommandText = l_sql
			cmExecute l_cm, l_sql, 0
		end if
		if l_tdnro = 2 then ' VACACION
			l_sql = "DELETE FROM lic_vacacion " 
			l_sql = l_sql & "WHERE emp_licnro = " & l_emp_licnro
			l_cm.activeconnection = Cn
			l_cm.CommandText = l_sql
			cmExecute l_cm, l_sql, 0
			
			l_sql = "DELETE FROM vacnotif " 
			l_sql = l_sql & "WHERE emp_licnro = " & l_emp_licnro
			l_cm.activeconnection = Cn
			l_cm.CommandText = l_sql
			cmExecute l_cm, l_sql, 0
		end if
		if (l_tdnro = 8) OR (l_tdnro = 32) then ' ENFERMEDAD
			
			l_sql = "SELECT * FROM lic_enfer " 
			l_sql = l_sql & "WHERE emp_licnro = " & l_emp_licnro
		    rsOpen l_rs, cn, l_sql, 0 
			
		    do until l_rs.eof
			   l_sql = "DELETE FROM lic_enfer " 
			   l_sql = l_sql & "WHERE emp_licnro = " & l_emp_licnro
			   l_sql = l_sql & "  AND vismednro  = " & l_rs("vismednro")
			   l_cm.activeconnection = Cn
			   l_cm.CommandText = l_sql
			   cmExecute l_cm, l_sql, 0
			   
		       l_rs.MoveNext
		    loop
			l_rs.close
		end if
		if l_tdnro = 9 or l_tdnro=13 or l_tdnro = 14  then ' accidente
			l_sql = "DELETE FROM lic_accid " 
			l_sql = l_sql & "WHERE emp_licnro = " & l_emp_licnro
			l_cm.activeconnection = Cn
			l_cm.CommandText = l_sql
			cmExecute l_cm, l_sql, 0
			
			l_sql = "SELECT * FROM licencia_visita " 
			l_sql = l_sql & "WHERE emp_licnro = " & l_emp_licnro
		    rsOpen l_rs, cn, l_sql, 0 
			
		    do until l_rs.eof
			   l_sql = "DELETE FROM licencia_visita " 
			   l_sql = l_sql & "WHERE emp_licnro = " & l_emp_licnro
			   l_sql = l_sql & "  AND visitamed  = " & l_rs("visitamed")
			   l_cm.activeconnection = Cn
			   l_cm.CommandText = l_sql
			   cmExecute l_cm, l_sql, 0
			   
		       l_rs.MoveNext
		    loop
			l_rs.close
		end if
		
		'set l_cm = Server.CreateObject("ADODB.Command")
		l_sql = "DELETE FROM emp_lic " 
		l_sql = l_sql & "WHERE emp_licnro = " & l_emp_licnro
		l_cm.activeconnection = Cn
		l_cm.CommandText = l_sql
		cmExecute l_cm, l_sql, 0
		
		
		l_sql = "DELETE FROM gti_justificacion WHERE jussigla = 'LIC' and juscodext = " & l_emp_licnro
		l_cm.CommandText = l_sql
		cmExecute l_cm, l_sql, 0
		
		
		l_sql = "DELETE FROM cysfirmas where cystipnro = 6 and cysfircodext = '" & l_emp_licnro & "' "
		l_cm.CommandText = l_sql
		cmExecute l_cm, l_sql, 0
		
		'Borro el campo licnrosig de las licencias que la tengan como siguiente.
		'l_sql = "SELECT emp_licnro, tdnro, empleado, licnrosig, elfechadesde, elfechahasta "
		'l_sql = l_sql   & " FROM emp_lic "
		'l_sql  = l_sql  & " WHERE licnrosig = " & l_emp_licnro
		
		'rsOpen l_rs, cn, l_sql, 0 
		
		'do until l_rs.eof
		'   	l_sql = "UPDATE emp_lic SET" 
		'   	l_sql = l_sql & " licnrosig = ''  "
		'   	l_sql = l_sql & " WHERE emp_licnro = " & l_rs("emp_licnro")
		'	
		'   	l_cm.activeconnection = Cn
		'   	l_cm.CommandText = l_sql
		'   	cmExecute l_cm, l_sql, 0
		'	
		'   	l_rs.MoveNext
		'loop
		'l_rs.close
		
		cn.CommitTrans
		
'		cn.Close

		Set cn = Nothing
		
		%>
		<script>
		  window.opener.ifrm.location.reload();
		  alert('Operación Realizada.');  
		  window.close();
		</script>
		
		<%
		
	end if
end if
%>

</body>
</html>

<% Option Explicit %>
<!--#include virtual="/serviciolocal/shared/db/conn_db.inc"-->
<!--#include virtual="/serviciolocal/shared/inc/const.inc"-->
<!--#include virtual="/serviciolocal/ess/shared/inc/accesoESS.inc"-->
<!--
Archivo: novedades_empleado_liq_04.asp
Descripción: abm de novedades del empleado
Autor: FFavre 
Fecha: 10-03
Modificado: 
	16-11-03 FFavre Se agrego firma
	26-10-05 - Leticia A. - Se comento lo de firmas para Autogestion.
	27-10-05 - Leticia A. - Se adecuo a Autogestion.
-->
<%
 Dim l_cm
 Dim l_sql
 Dim l_rs
 Dim l_nenro
 Dim l_tipAutorizacion  'Es el tipo del circuito de firmas
 Dim l_HayAutorizacion  'Es para ver si las autorizaciones estan activas
 Dim l_PuedeVer         'Es para ver si las autorizaciones estan activas
 Dim l_ternro 
 
 l_nenro  = request("nenro")
  l_nenro  	= request.querystring("nenro")
 
 Set l_rs = Server.CreateObject("ADODB.RecordSet")
 Set l_cm = Server.CreateObject("ADODB.Command")
 
 l_ternro = l_ess_ternro
 if l_ternro = "" or l_ternro= 0 then
 	response.end
 end if
 
 l_tipAutorizacion = 5
 
 'l_sql = "select cystipo.* from cystipo "
 'l_sql = l_sql & "where (cystipo.cystipact = -1) and cystipo.cystipnro = " & l_tipAutorizacion
 
 'rsOpen l_rs, cn, l_sql, 0 
 'l_HayAutorizacion = not l_rs.eof
 'l_rs.close
 
 cn.beginTrans
 if l_HayAutorizacion then
	'l_sql = "select cysfirautoriza, cysfirsecuencia, cysfirdestino from cysfirmas "
	'l_sql = l_sql & "where cysfirmas.cystipnro = " & l_tipAutorizacion & " and cysfirmas.cysfircodext = '" & l_nenro & "' " 
	'l_sql = l_sql & "order by cysfirsecuencia desc"
	'rsOpen l_rs, cn, l_sql, 0 
 	
	'l_PuedeVer = False
 	
	'if not l_rs.eof then
 		'if (l_rs("cysfirautoriza") = session("UserName")) or (l_rs("cysfirdestino") = session("UserName")) then 
	   		'Es una modificación del ultimo o es el nuevo que autoriza 
    		'l_PuedeVer = True 
    	'end if
 	'end if
	'l_rs.close
 	'If not l_PuedeVer then
    	'response.write "<script>alert('No esta autorizado a ver o modificar este registro.');window.close()</script>"
		'response.end
 	'End if
	
	'l_sql = "DELETE FROM cysfirmas where cystipnro = " & l_tipAutorizacion & " and cysfircodext = '" & l_nenro & "' "
 	'l_cm.CommandText = l_sql
 	'l_cm.activeconnection = Cn
 	'cmExecute l_cm, l_sql, 0
 End if
 
 l_sql = "DELETE FROM novemp "
 l_sql = l_sql & "WHERE nenro = " & l_nenro
 l_cm.activeconnection = Cn
 l_cm.CommandText = l_sql
 cmExecute l_cm, l_sql, 0
 
 cn.CommitTrans
 
 cn.Close
 Set cn = Nothing
 Response.write "<script>alert('Operación Realizada.');window.opener.ifrm.location.reload();window.close();</script>"
%>

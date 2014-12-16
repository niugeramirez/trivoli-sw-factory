<% Option Explicit %>
<!--#include virtual="/serviciolocal/shared/inc/sec.inc"-->
<!--#include virtual="/serviciolocal/shared/inc/const.inc"-->
<!--#include virtual="/serviciolocal/shared/db/conn_db.inc"-->
<!--#include virtual="/serviciolocal/shared/inc/fecha.inc"-->
<% 
'Archivo: contracts_con_03.asp
'Descripción: ABM de Contracts
'Autor : Raul Chinestra
'Fecha: 28/11/2007

on error goto 0

Dim l_tipo
Dim l_cm
Dim l_sql

Dim l_legpar1
Dim l_legpar2
Dim l_legpar3
Dim l_legnro
Dim l_legape
Dim l_legnom
Dim l_legfecing
Dim l_legdni
Dim l_legfecnac 
Dim l_legdom
Dim l_legtel
Dim l_pronro
Dim l_legins
Dim l_leginsedu
Dim l_legcobsoc
Dim l_legabo
Dim l_mednro


l_tipo 		  = request.querystring("tipo")
l_legpar1     = request.Form("legpar1")
l_legpar2     = request.Form("legpar2")
l_legpar3     = request.Form("legpar3")
l_legnro      = request.Form("legnro")
l_legape      = request.Form("legape")
l_legnom      = request.Form("legnom")
l_legfecing   = request.Form("legfecing")
l_legdni 	  = request.Form("legdni")
l_legfecnac   = request.Form("legfecnac")
l_legdom	  = request.Form("legdom")
l_legtel	  = request.Form("legtel")
l_pronro	  = request.Form("pronro")
l_legins	  = request.Form("legins")
l_leginsedu   = request.Form("leginsedu")
l_legcobsoc   = request.Form("legcobsoc")
l_legabo	  = request.Form("legabo")
l_mednro	  = request.Form("mednro")


if len(l_legfecing) = 0 then
	l_legfecing = "null"
else 
	l_legfecing = cambiafecha(l_legfecing,"YMD",true)	
end if 
if len(l_legfecnac) = 0 then
	l_legfecnac = "null"
else 
	l_legfecnac = cambiafecha(l_legfecnac,"YMD",true)	
end if 

set l_cm = Server.CreateObject("ADODB.Command")
if l_tipo = "A" then 
	l_sql = "INSERT INTO ser_legajo "
	l_sql = l_sql & " (legpar1, legpar2, legpar3,legape,legnom, legfecing,legdni,legfecnac, legdom, legtel,  pronro, legins, l_leginsedu,legcobsoc, legabo, mednro)"
	l_sql = l_sql & " VALUES ('" & l_legpar1 & "','" & l_legpar2 & "','" & l_legpar3 & "','" & l_legape & "','" & l_legnom & "'," & l_legfecing & ",'" & l_legdni & "'," & l_legfecnac & ",'" & l_legdom & "','" & l_legtel & "'," & l_pronro & ",'" & l_legins & "','" & l_leginsedu & "','" & l_legcobsoc & "','" & l_legabo & "'," & l_mednro & ")"
else
	l_sql = "UPDATE ser_legajo "
	l_sql = l_sql & " SET legpar1    = '" & l_legpar1 & "'"
	l_sql = l_sql & "    ,legpar2    = '" & l_legpar2 & "'"
	l_sql = l_sql & "    ,legpar3    = '" & l_legpar3 & "'"
	l_sql = l_sql & "    ,legape     = '" & l_legape & "'"
	l_sql = l_sql & "    ,legnom     = '" & l_legnom & "'"
	l_sql = l_sql & "    ,legfecing  = " & l_legfecing
	l_sql = l_sql & "    ,legdni     = '" & l_legdni & "'"
	l_sql = l_sql & "    ,legfecnac  = " & l_legfecnac
	l_sql = l_sql & "    ,legdom     = '" & l_legdom & "'"
	l_sql = l_sql & "    ,legtel     = '" & l_legtel & "'"	
	l_sql = l_sql & "    ,pronro     =  " & l_pronro
	l_sql = l_sql & "    ,legins     = '" & l_legins & "'"
	l_sql = l_sql & "    ,leginsedu  = '" & l_leginsedu & "'"	
	l_sql = l_sql & "    ,legcobsoc  = '" & l_legcobsoc & "'"	
	l_sql = l_sql & "    ,legabo     = '" & l_legabo & "'"
	l_sql = l_sql & "    ,mednro      = " & l_mednro
	l_sql = l_sql & " WHERE legnro = " & l_legnro
end if
response.write l_sql & "<br>"
l_cm.activeconnection = Cn
l_cm.CommandText = l_sql
cmExecute l_cm, l_sql, 0
Set l_cm = Nothing

Response.write "<script>alert('Operación Realizada.');window.parent.opener.ifrm.location.reload();window.parent.close();</script>"
%>


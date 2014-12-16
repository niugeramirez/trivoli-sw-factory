<% Option Explicit %>
<!--#include virtual="/serviciolocal/shared/inc/sec.inc"-->
<!--#include virtual="/serviciolocal/shared/inc/const.inc"-->
<!--#include virtual="/serviciolocal/shared/inc/const_eva.inc"-->
<!--#include virtual="/serviciolocal/shared/inc/sqls.inc"-->
<!--#include virtual="/serviciolocal/shared/db/conn_db.inc"-->
<!--#include virtual="/serviciolocal/shared/inc/fecha.inc"-->
<script>
//var xc = screen.availWidth;
//var yc = screen.availHeight;
//window.moveTo(xc,yc);	
//window.resizeTo(150,150);
</script>
<script src="/serviciolocal/shared/js/fn_ayuda.js"></script>
<script src="/serviciolocal/shared/js/fn_windows.js"></script>
<script src="/serviciolocal/shared/js/fn_fechas.js"></script>
<%
'================================================================================
'Archivo		: proyecto_eva_ag_03.asp
'Descripción	: Grabar Proyecto
'Autor			: CCRossi
'Fecha			: 31-08-2004
'Modificado		: 11-03-2005 - LAmadio - enviar mail al gerente del proyecto (si lo da de alta otra persona)
' 				: 14-05-2005 - update proyrevisor 
'				: 26-07-2005 - considerar cuando se puede o no modificar el revisor
'				: 28-07-2005 - crear el evento asociado al proyecto 
'==================================================================================

on error goto 0

Dim l_tipo
Dim l_rs
Dim l_cm
Dim l_sql

Dim l_evaproynro
Dim l_ternro
Dim l_evaproynom
Dim l_evaproydext
Dim l_evaproyfdd
Dim l_evaproyfht
Dim l_proysocio		
Dim l_proygerente	
Dim l_proyrevisor	
Dim l_proyaux1
Dim l_proyaux2
Dim l_estrnro
Dim l_perfil
Dim l_evaengnro
Dim l_evapernro
dim l_revanterior

l_tipo 		= request("tipo")
l_perfil	= request.Form("perfil")

l_evaproynro	= request.Form("evaproynro")
l_evaproynom 	= request.Form("evaproynom")
l_evaproydext 	= request.Form("evaproydext")
l_ternro		= request.Form("ternro")
l_evaproyfdd	= request.Form("evaproyfdd")
l_evaproyfht	= request.Form("evaproyfht")
	' OBS: viene como dato el EMPLEG y no el TERNRO -------- 
l_proysocio		= request.Form("proysocio") ' viene el empleg!
l_proygerente	= request.Form("proygerente") ' viene el empleg!
l_proyrevisor	= request.Form("proyrevisor") ' viene el empleg!
l_proyaux1		= request.Form("proyaux1") ' viene el empleg!
l_proyaux2		= request.Form("proyaux2") ' viene el empleg!
l_estrnro		= request.Form("estrnro")
l_evaengnro		= Request.Form("evaengnro")
l_evapernro		= Request.Form("evapernro")



Set l_rs = Server.CreateObject("ADODB.RecordSet")

' 
if l_evaproydext <> "" then
	l_evaproynom = left(l_evaproydext,30)
end if

	' BUSCO LOS TERNRO DE proysocio - proygerente -proyrevisor - ....
if trim(l_proysocio)<>"" and not isnull(l_proysocio) then
	l_sql = "SELECT ternro  FROM empleado  WHERE empleg = "& l_proysocio
	rsOpen l_rs, cn, l_sql, 0 
	if not l_rs.eof then
		l_proysocio= l_rs("ternro")
	end if
	l_rs.close
else
	l_proysocio = ""
end if
if trim(l_proygerente)<>"" and not isnull(l_proygerente) then
	l_sql = "SELECT ternro  FROM empleado WHERE empleg ="& l_proygerente
	rsOpen l_rs, cn, l_sql, 0 
	if not l_rs.eof then
		l_proygerente= l_rs("ternro")
	end if
	l_rs.close
else
	l_proygerente= ""
end if

if trim(l_proyrevisor)<>"" and not isnull(l_proyrevisor) then
	l_sql = "SELECT ternro  FROM empleado WHERE empleg = "& l_proyrevisor
	rsOpen l_rs, cn, l_sql, 0 
	if not l_rs.eof then
		l_proyrevisor= l_rs("ternro")
	end if
	l_rs.close
else
	l_proyrevisor=""
end if

if trim(l_proyaux1)<>"" and not isnull(l_proyaux1) then
	l_sql = "SELECT ternro  FROM empleado WHERE empleg = "& l_proyaux1
	rsOpen l_rs, cn, l_sql, 0 
	if not l_rs.eof then
		l_proyaux1= l_rs("ternro")
	end if
	l_rs.close
else
	l_proyaux1= ""
end if
if trim(l_proyaux2)<>"" and not isnull(l_proyaux2) then
	l_sql = "SELECT ternro  FROM empleado WHERE empleg = "& l_proyaux2
	rsOpen l_rs, cn, l_sql, 0 
	if not l_rs.eof then
		l_proyaux2= l_rs("ternro")
	end if
	l_rs.close
else
	l_proyaux2 = ""
end if

if trim(l_evaproyfdd)<>"" then
	l_evaproyfdd = cambiafecha(l_evaproyfdd,"","")
else
	l_evaproyfdd = "NULL"
end if
if trim(l_evaproyfht)<>"" then
	l_evaproyfht = cambiafecha(l_evaproyfht,"","")
else
	l_evaproyfht = "NULL"
end if



if l_tipo = "A" then 
	set l_cm = Server.CreateObject("ADODB.Command")
	l_sql = "INSERT INTO evaproyecto "
	l_sql = l_sql & " (evaproynom,evaproydext,evaproyfdd,evaproyfht,proysocio,proygerente,proyrevisor "
	if l_proyaux1 <> "" then
		l_sql = l_sql & " ,proyaux1 "
	end if
	if l_proyaux2 <> "" then
		l_sql = l_sql & " ,proyaux2 "
	end if
	l_sql = l_sql & " ,estrnro,evapernro,evaengnro) "
	l_sql = l_sql & " VALUES ('"& l_evaproynom &"', '"
	l_sql = l_sql & l_evaproydext  & "',"
	l_sql = l_sql & l_evaproyfdd  & ","
	l_sql = l_sql & l_evaproyfht  & ","
	l_sql = l_sql & l_proysocio   & ","
	l_sql = l_sql & l_proygerente & ","
	l_sql = l_sql & l_proyrevisor & ","
	if l_proyaux1 <> "" then
		l_sql =l_sql & l_proyaux1 & ","
	end if
	if l_proyaux2 <> "" then
		l_sql =l_sql & l_proyaux2 & ","
	end if
	l_sql = l_sql & l_estrnro	& ","
	l_sql = l_sql & l_evapernro	& ","
	l_sql = l_sql & l_evaengnro & ")"
	l_cm.activeconnection = Cn
	l_cm.CommandText = l_sql
	cmExecute l_cm, l_sql, 0
	Set l_cm = Nothing
	
	Set l_rs = Server.CreateObject("ADODB.RecordSet")
	l_sql = fsql_seqvalue("evaproynro","evaproyecto") 
	rsOpen l_rs, cn, l_sql, 0 
	l_evaproynro=l_rs("evaproynro") 
	l_rs.Close 
	Set l_rs = Nothing
	
	set l_cm = Server.CreateObject("ADODB.Command")
	l_sql = "INSERT INTO evaproyemp (evaproynro,ternro) "
	l_sql = l_sql & " VALUES (" & l_evaproynro &","& l_ternro & ")"
	l_cm.activeconnection = Cn 
	l_cm.CommandText = l_sql 
	cmExecute l_cm, l_sql, 0 
	Set l_cm = Nothing 
	
	'crear el EVENTO  asociado 
	set l_cm = Server.CreateObject("ADODB.Command") 
	l_sql = "INSERT INTO evaevento (evaproynro, evaperact, evaevedesabr, evaevefdesde, evaevefhasta)"
	l_sql = l_sql & " VALUES ("& l_evaproynro & ","& l_evapernro &",'" & l_evaproynom & "',"
	l_sql = l_sql & l_evaproyfdd &"," & l_evaproyfht &")"
	l_cm.activeconnection = Cn 
	l_cm.CommandText = l_sql 
	cmExecute l_cm, l_sql, 0 
	Set l_cm = Nothing
	
else
	
	set l_cm = Server.CreateObject("ADODB.Command")
	l_sql = "UPDATE evaproyecto SET "
	l_sql = l_sql & " evaproynom	= '" & l_evaproynom  & "',"
	l_sql = l_sql & " evaproydext	= '" & l_evaproydext & "',"
	if l_proyrevisor <> "" then
		l_sql = l_sql & " proyrevisor = " & l_proyrevisor & ","
	end if
	l_sql = l_sql & " evapernro     = " & l_evapernro    & ","
	l_sql = l_sql & " evaproyfdd    = " & l_evaproyfdd   & ","
	l_sql = l_sql & " evaproyfht    = " & l_evaproyfht   & " "
	l_sql = l_sql & " WHERE evaproynro = " & l_evaproynro
	l_cm.activeconnection = Cn
	l_cm.CommandText = l_sql
	cmExecute l_cm, l_sql, 0
	Set l_cm = Nothing
	
	' actualizar el periodo del evento del proyecto y el nombre del evento
	set l_cm = Server.CreateObject("ADODB.Command")
	l_sql = "UPDATE evaevento SET "
	l_sql = l_sql & " evaevedesabr	= '" & l_evaproynom & "',"
	l_sql = l_sql & " evaperact     = " & l_evapernro   & ","
	l_sql = l_sql & " evaevefdesde  = " & l_evaproyfdd  & ","
	l_sql = l_sql & " evaevefhasta  = " & l_evaproyfht  & " "
	l_sql = l_sql & " WHERE evaproynro = " & l_evaproynro
	l_cm.activeconnection = Cn
	l_cm.CommandText = l_sql
	cmExecute l_cm, l_sql, 0
	Set l_cm = Nothing
	
end if


'if l_perfil="empleado" and l_tipo="A" then 
	' HAY QUE ENVIAR MAIL - EL EMPLEADO AL GERENTE DEL PROYECTO - (no ver si el empl tiene estrc gerente)
if l_tipo="A" then 
	if (trim(l_ternro) <>  trim(l_proygerente)) then ' el que dio de alta proyecto no es el gerente del proyecto 
		if cUsaMail= -1 and trim(l_ternro)<>"" and trim(l_proygerente)<>"" then %>
		<script>
		abrirVentanaH('mailaviso_eva_00.asp?ternro=<%=l_ternro%>&proygerente=<%=l_proygerente%>&evaproynro=<%=l_evaproynro%>',"",5,5);
		</script>
<%		end if 
	end if 
end if


cn.Close
Set cn = Nothing

Response.write "<script>alert('Operación Realizada.');window.opener.ifrm.location.reload();window.close();</script>"
%>

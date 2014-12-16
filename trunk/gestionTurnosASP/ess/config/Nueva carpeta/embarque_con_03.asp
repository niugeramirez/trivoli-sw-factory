<% Option Explicit %>
<!--#include virtual="/serviciolocal/shared/inc/sec.inc"-->
<!--#include virtual="/serviciolocal/shared/inc/const.inc"-->
<!--#include virtual="/serviciolocal/shared/db/conn_db.inc"-->
<% 
'Archivo: embarque_con_03.asp
'Descripción: ABM de embarque
'Autor : Gustavo Manfrin
'Fecha: 18/09/2006

Dim l_tipo
Dim l_cm
Dim l_sql
Dim l_rs

Dim l_embnro
Dim l_embcod
Dim l_embkiltot
Dim l_embsim
Dim l_embact
Dim l_desnro
Dim l_cornro
Dim l_depnro
Dim l_entnro
Dim l_ordnro
Dim l_connro
  

l_tipo 		= request.querystring("tipo")
l_embnro 	= request.Form("embnro")
l_embcod 	= request.Form("embcod")
l_embkiltot	= request.Form("embkiltot")
l_embsim	= request.Form("embsim")
l_embact	= request.Form("embact")
l_desnro	= request.Form("desnro")
l_cornro	= request.Form("cornro")
l_depnro	= request.Form("depnro")
'l_entnro	= request.Form("entnro")
l_entnro = "null"
l_ordnro	= request.Form("ordnro")
l_connro	= request.Form("connro")


if isnull(l_connro) then
	response.write "SI"
else 	
	response.write "NO"
end if  

'response.write "------" & l_connro
'response.end

if len(l_embact)>0 then
	l_embact = -1
else
	l_embact = 0
end if


	set l_rs = Server.CreateObject("ADODB.RecordSet")
	
	' Obtengo el codigo interno del destinatario
	' Si el destinatario viene con un valor busco su codigo interno
	if l_desnro <> "" then 
		l_sql = "SELECT vencornro"
		l_sql  = l_sql  & " FROM tkt_vencor "
		l_sql  = l_sql  & " WHERE tkt_vencor.vencorcod = '" & l_desnro & "'"
		rsOpen l_rs, cn, l_sql, 0
		if not l_rs.eof  then 
			l_desnro= l_rs("vencornro") 
		else
			l_desnro="null"
		end if 		
		l_rs.Close
	end if

	set l_cm = Server.CreateObject("ADODB.Command")
	l_cm.activeconnection = Cn

	
	if l_tipo = "A" then 
		l_sql = "INSERT INTO tkt_embarque "
		l_sql = l_sql & " (embcod, embkiltot, embkilemb, embsim, embact, "
		l_sql = l_sql & "  desnro, cornro, depnro, entnro, ordnro, connro)"
		l_sql = l_sql & " VALUES (" & l_embcod 
		l_sql = l_sql & "," & l_embkiltot
		l_sql = l_sql & "," & 0
		l_sql = l_sql & ",'" & l_embsim & "'"
		l_sql = l_sql & "," & l_embact
		l_sql = l_sql & "," & l_desnro
		l_sql = l_sql & "," & l_cornro
		l_sql = l_sql & "," & l_depnro
		l_sql = l_sql & "," & l_entnro
		l_sql = l_sql & "," & l_ordnro
		if l_connro = 0 then
		   l_sql = l_sql & ", null "
		else
		   l_sql = l_sql & "," & l_connro
		end if   
		l_sql = l_sql & ")"
	else
		l_sql = "UPDATE tkt_embarque "
		l_sql = l_sql & " SET embcod = " & l_embcod
		l_sql = l_sql & ", embkiltot = " & l_embkiltot
		l_sql = l_sql & ", embsim = '" & l_embsim & "'"
		l_sql = l_sql & ", embact = " & l_embact
		l_sql = l_sql & ", desnro = " & l_desnro
		l_sql = l_sql & ", cornro = " & l_cornro
		l_sql = l_sql & ", depnro = " & l_depnro
		l_sql = l_sql & ", entnro = " & l_entnro
		l_sql = l_sql & ", ordnro = " & l_ordnro
		if l_connro = 0 then
  		  l_sql = l_sql & ", connro = null "
		else
 		  l_sql = l_sql & ", connro = " & l_connro
		end if  
		l_sql = l_sql & " WHERE  embnro = " & l_embnro
	end if
	l_cm.CommandText = l_sql
	cmExecute l_cm, l_sql, 0
	
	Set l_cm = Nothing
	Response.write "<script>alert('Operación Realizada.');window.parent.opener.ifrm.location.reload();window.parent.close();</script>"
%>


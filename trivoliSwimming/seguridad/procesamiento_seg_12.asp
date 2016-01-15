<% Option Explicit %>
<!--#include virtual="/trivoliSwimming/shared/db/conn_db.inc"-->
<!--#include virtual="/trivoliSwimming/shared/inc/const.inc"-->
<!--#include virtual="/trivoliSwimming/shared/inc/fecha.inc"-->
<% 
'Archivo: procesamiento_seg_12.asp
'Descripción:
'Autor : Lisandro Moro
'Fecha : 10/03/2005
'Modificado:

on error goto 0

Dim l_cm
Dim l_sql

Dim l_bprcestado
Dim l_bpronro
Dim l_bprcfecdesde
Dim l_bprcfechasta
Dim l_bprcUrgente
Dim l_bprcterminar
Dim l_estadopasar
Dim l_bprcparam
Dim l_btprcnro
Dim l_arr
Dim i

l_bpronro = request.Form("bpronro")
l_bprcfecdesde = request.Form("bprcfecdesde")
l_bprcfechasta = request.Form("bprcfechasta")
l_bprcUrgente = request.Form("bprcurgente")
l_bprcterminar = request.Form("bprcterminar")
l_bprcestado = request.Form("bprcestado")
l_bprcparam = request.Form("bprcparam")
l_btprcnro = request.Form("btprcnro")

if l_bprcUrgente = "on" then
	l_bprcUrgente = "-1"
else
	l_bprcUrgente = "0"
end if

if l_bprcterminar = "on" then
	l_bprcterminar = "-1"
else
	l_bprcterminar = "0"
end if


set l_cm = Server.CreateObject("ADODB.Command")

if len(l_bprcfecdesde)>0 then
	l_sql = "UPDATE batch_proceso set bprcfecdesde=" & cambiafecha(l_bprcfecdesde,"YMD",true)
else
	l_sql = "UPDATE batch_proceso set bprcfecdesde=null" 
end if

if len(l_bprcfechasta)>0 then
	l_sql = l_sql & ", bprcfechasta=" & cambiafecha(l_bprcfechasta,"YMD",true)
else
	l_sql = l_sql & ", bprcfechasta= null" 
end if

l_sql = l_sql & ", bprcurgente=" & l_bprcurgente
l_sql = l_sql & ", bprcterminar=" & l_bprcterminar

if l_bprcterminar then
	l_sql = l_sql & ", bprcestado= 'Abortando' "
else
	if l_bprcestado <> "Pendiente" then
		l_estadopasar = request.Form("estadopasar")
		if l_estadopasar = "on" then
		   l_sql = l_sql & ", bprcestado= 'Pendiente' "
		   'Si es un proceso de liquidacion tengo que poner el quinto parametro de bprcparam en cero
		   if CInt(l_btprcnro) = 3 then
		      l_arr = split(l_bprcparam,".")
			  l_bprcparam = ""
			  for i = 0 to UBound(l_arr)
			     if l_bprcparam = "" then
				    l_bprcparam = l_arr(i)
				 else
				    if i = 4 then
					   l_bprcparam = l_bprcparam & ".0" 
					else
					   l_bprcparam = l_bprcparam & "." & l_arr(i)					
					end if
				 end if
			  next
		      l_sql = l_sql & ", bprcparam= '" & l_bprcparam & "'"			  
		   end if 
		end if
	end if
end if

l_sql = l_sql & " WHERE bpronro=" & l_bpronro
'response.write l_sql
l_cm.activeconnection = Cn
l_cm.CommandText = l_sql
cmExecute l_cm, l_sql, 0

cn.Close
set cn = Nothing
Response.write "<script>alert('Operación realizada');window.opener.location.reload();window.close();</script>"
%>


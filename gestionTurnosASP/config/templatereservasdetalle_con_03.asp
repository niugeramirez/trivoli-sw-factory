<% Option Explicit %>
<!--#include virtual="/turnos/shared/inc/sec.inc"-->
<!--#include virtual="/turnos/shared/inc/const.inc"-->
<!--#include virtual="/turnos/shared/db/conn_db.inc"-->
<% 
'Archivo: companies_con_03.asp
'Descripción: ABM de Companies
'Autor : Raul Chinestra
'Fecha: 26/11/2007

Dim l_tipo
Dim l_cm
Dim l_sql

Dim l_id
Dim l_titulo
Dim l_descripcion

Dim l_idtemplatereserva
Dim l_calhordes1
Dim l_calhordes2
Dim l_calhorhas1
Dim l_calhorhas2

Dim l_calhordes
dIM l_calhorhas
Dim l_intervaloTurnoMinutos


Dim l_lu
Dim l_ma
Dim l_mi
Dim l_ju
Dim l_vi
Dim l_sa
Dim l_do




l_tipo 		  = request.querystring("tipo")
l_id 	      = request.Form("id")
l_titulo	  = request.Form("titulo")
l_descripcion = request.Form("descripcion")

l_idtemplatereserva = request("idtemplatereserva")
l_calhordes1 = request("calhordes1")
l_calhordes2 = request("calhordes2")
l_calhorhas1 = request("calhorhas1")
l_calhorhas2 = request("calhorhas2")
l_intervaloTurnoMinutos = request("intervaloTurnoMinutos")
l_lu            = request.Form("lu")
l_ma            = request.Form("ma")
l_mi            = request.Form("mi")
l_ju            = request.Form("ju")
l_vi            = request.Form("vi")
l_sa            = request.Form("sa")
l_do            = request.Form("do")

if l_lu = "on" then
l_lu = "S"
else
l_lu = "N"
end if

if l_ma = "on" then
l_ma = "S"
else
l_ma = "N"
end if

if l_mi = "on" then
l_mi = "S"
else
l_mi = "N"
end if

if l_ju = "on" then
l_ju = "S"
else
l_ju = "N"
end if

if l_vi = "on" then
l_vi = "S"
else
l_vi = "N"
end if

if l_sa = "on" then
l_sa = "S"
else
l_sa = "N"
end if

if l_do = "on" then
l_do = "S"
else
l_do = "N"
end if

l_calhordes = l_calhordes1 & ":" &  l_calhordes2
l_calhorhas = l_calhorhas1 & ":" &  l_calhorhas2

set l_cm = Server.CreateObject("ADODB.Command")
if l_tipo = "A" then 
	l_sql = "INSERT INTO templatereservasdetalleresumido "
	l_sql = l_sql & " (titulo, idtemplatereserva, horaInicial, horaFinal, intervaloTurnoMinutos, dia1, dia2, dia3, dia4, dia5, dia6, dia7,empnro,created_by,creation_date,last_updated_by,last_update_date )"
	l_sql = l_sql & " VALUES ('" & l_titulo & "'," & l_idtemplatereserva & ",'" & l_calhordes & "','" & l_calhorhas & "','" & l_intervaloTurnoMinutos & "','" & l_do & "','" & l_lu & "','" & l_ma & "','" & l_mi & "','" & l_ju & "','" & l_vi & "','" & l_sa & "','" & session("empnro") & "','"&session("loguinUser")&"',GETDATE(),'"&session("loguinUser")&"',GETDATE())"
else
	l_sql = "UPDATE templatereservasdetalleresumido "
	l_sql = l_sql & " SET titulo = '" & l_titulo & "'"
	l_sql = l_sql & " , horaInicial = '" & l_calhordes & "'"
	l_sql = l_sql & " , horafinal = '" & l_calhorhas & "'"
	l_sql = l_sql & " , intervaloTurnoMinutos = '" & l_intervaloTurnoMinutos & "'"
	l_sql = l_sql & " , dia1 = '" & l_do & "'"
	l_sql = l_sql & " , dia2 = '" & l_lu & "'"
	l_sql = l_sql & " , dia3 = '" & l_ma & "'"
	l_sql = l_sql & " , dia4 = '" & l_mi & "'"
	l_sql = l_sql & " , dia5 = '" & l_ju & "'"
	l_sql = l_sql & " , dia6 = '" & l_vi & "'"
	l_sql = l_sql & " , dia7 = '" & l_sa & "'"	
	l_sql = l_sql & "    ,last_updated_by = '" &session("loguinUser") & "'"
	l_sql = l_sql & "    ,last_update_date = GETDATE()" 
	l_sql = l_sql & " WHERE id = " & l_id
end if
'response.write l_sql & "<br>"
l_cm.activeconnection = Cn
l_cm.CommandText = l_sql
cmExecute l_cm, l_sql, 0
Set l_cm = Nothing

Response.write "<script>alert('Operación Realizada.');window.parent.opener.ifrm.location.reload();window.parent.close();</script>"
%>


<% Option Explicit %>
<!--#include virtual="/turnos/shared/inc/sec.inc"-->
<!--#include virtual="/turnos/shared/inc/const.inc"-->
<!--#include virtual="/turnos/shared/db/conn_db.inc"-->
<!--#include virtual="/turnos/shared/inc/fecha.inc"-->
<% 

Dim l_tipo
Dim l_cm
Dim l_sql
Dim l_rs

Dim l_id
Dim l_calfec
Dim l_descripcion


Dim l_calhordes1
Dim l_calhordes2
Dim l_calhorhas1
Dim l_calhorhas2

Dim l_calhordes
dIM l_calhorhas
Dim l_intervaloTurnoMinutos

Dim l_hora
Dim l_horafin



l_tipo 		  = request.querystring("tipo")
l_id 	      = request.Form("id")

l_calfec = request("calfec")

l_calhordes1 = request("calhordes1")
l_calhordes2 = request("calhordes2")
l_calhorhas1 = request("calhorhas1")
l_calhorhas2 = request("calhorhas2")
l_intervaloTurnoMinutos = request("intervaloTurnoMinutos")

l_calhordes = l_calhordes1 & ":" &  l_calhordes2  & ":00"
l_calhorhas = l_calhorhas1 & ":" &  l_calhorhas2  & ":00"


Set l_rs = Server.CreateObject("ADODB.RecordSet")


l_hora =  l_calhordes
l_horafin = DateAdd("n", cint(l_intervaloTurnoMinutos) , l_hora)
							
'response.write "l_hora" & l_hora
'response.write "l_horafin" & l_horafin
'response.end
set l_cm = Server.CreateObject("ADODB.Command")
Do While DATEDIFF("n", cdate( l_horafin ), cdate( l_calhorhas )) >= 0

	l_sql = "INSERT INTO calendarios "
	l_sql = l_sql & "(fechahorainicio, fechahorafin, estado, idrecursoreservable, tipo  ) "
	l_sql = l_sql & "VALUES (" & cambiaformato (l_calfec,l_hora )  & "," & cambiaformato (l_calfec,l_horafin )  & ",'ACTIVO'," & l_id & ",'MANUAL')"
	l_cm.activeconnection = Cn
	l_cm.CommandText = l_sql
	cmExecute l_cm, l_sql, 0

	l_hora = DateAdd("n", clng(l_intervaloTurnoMinutos), l_hora)
	l_horafin = DateAdd("n", clng(l_intervaloTurnoMinutos), l_hora)
Loop




'set l_cm = Server.CreateObject("ADODB.Command")
'if l_tipo = "A" then 
'	l_sql = "INSERT INTO calendarios "
'	l_sql = l_sql & " (fechahoraInicio, fechahoraFin, idrecursoreservable, estado )"
'	l_sql = l_sql & " VALUES (" & cambiaformato (l_calfec,l_calhordes )  & "," & cambiaformato (l_calfec,l_calhorhas )  & "," & l_id & ",'ACTIVO'" & ")"

'end if
'response.write l_sql & "<br>"
'l_cm.activeconnection = Cn
'l_cm.CommandText = l_sql
'cmExecute l_cm, l_sql, 0
Set l_cm = Nothing

Response.write "<script>alert('Operación Realizada.');window.parent.opener.ifrm.location.reload();window.parent.close();</script>"
%>


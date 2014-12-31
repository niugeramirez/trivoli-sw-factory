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



l_tipo 		  = request.querystring("tipo")
l_id 	      = request.Form("id")

l_calfec = request("calfec")

l_calhordes1 = request("calhordes1")
l_calhordes2 = request("calhordes2")
l_calhorhas1 = request("calhorhas1")
l_calhorhas2 = request("calhorhas2")
l_intervaloTurnoMinutos = request("intervaloTurnoMinutos")

l_calhordes = l_calhordes1 & ":" &  l_calhordes2
l_calhorhas = l_calhorhas1 & ":" &  l_calhorhas2


Set l_rs = Server.CreateObject("ADODB.RecordSet")

'Verifico que no este repetida la descripción o el código externo
l_sql = "SELECT * "
l_sql = l_sql & " FROM calendarios "
l_sql = l_sql & " WHERE id=" & l_id
'l_sql = l_sql & " AND counro <> " & l_counro

rsOpen l_rs, cn, l_sql, 0
if not l_rs.eof then
    texto =  "Ya existe otro Country con esa Descripción."
end if 
l_rs.close




set l_cm = Server.CreateObject("ADODB.Command")
if l_tipo = "A" then 
	l_sql = "INSERT INTO calendarios "
	l_sql = l_sql & " (fechahoraInicio, fechahoraFin, idrecursoreservable, estado )"
	l_sql = l_sql & " VALUES (" & cambiaformato (l_calfec,l_calhordes )  & "," & cambiaformato (l_calfec,l_calhorhas )  & "," & l_id & ",'ACTIVO'" & ")"

end if
'response.write l_sql & "<br>"
l_cm.activeconnection = Cn
l_cm.CommandText = l_sql
cmExecute l_cm, l_sql, 0
Set l_cm = Nothing

Response.write "<script>alert('Operación Realizada.');window.parent.opener.ifrm.location.reload();window.parent.close();</script>"
%>


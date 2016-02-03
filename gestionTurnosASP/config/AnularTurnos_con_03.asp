<% Option Explicit %>
<!--#include virtual="/turnos/shared/inc/sec.inc"-->
<!--#include virtual="/turnos/shared/inc/const.inc"-->
<!--#include virtual="/turnos/shared/db/conn_db.inc"-->
<!--#include virtual="/turnos/shared/inc/fecha.inc"-->
<% 


on error goto 0

Dim l_tipo
Dim l_cm
Dim l_sql
Dim l_rs

dim l_id
dim l_motivo
dim l_estado
dim l_cantturnossimult  
dim l_cantsobreturnos     

Dim l_hd
Dim l_md
Dim l_hh
Dim l_mh

Dim l_opc

Dim	l_evento
Dim l_responsable 
Dim l_comentario 

Dim l_horadesde
Dim l_horahasta
Dim l_fechadesde
Dim l_fechahasta
Dim l_idrecursoreservable
Dim l_fecha
Dim l_hora


l_tipo 		         = request.querystring("tipo")
l_id                 = request.Form("id")
l_fechadesde         = request("fechadesde")
l_fechahasta       	 = request("fechahasta")
l_opc 				 = request("rbopc")
l_hd			     = request("hd") 
l_md			     = request("md")
l_hh			     = request("hh")
l_mh			     = request("mh")
l_motivo             = request("motivo")
l_idrecursoreservable = request.Form("idrecursoreservable")

' l_domicilio      = request.Form("domicilio")
'l_idobrasocial      = request.Form("legape")

l_horadesde = l_hd & ":" & l_md
l_horahasta = l_hh & ":" & l_mh


if l_tipo = "B" then
	l_estado = "ANULADO"
else
	l_estado = "ACTIVO"
end if

set l_cm = Server.CreateObject("ADODB.Command")
Set l_rs = Server.CreateObject("ADODB.RecordSet")

' Response.write "<script>alert('Operación " &l_opc&" Realizada.');</script>"

if l_opc = 1 then

	l_sql = "UPDATE calendarios "
	l_sql = l_sql & " SET motivo    = '" & l_motivo & "'"
	l_sql = l_sql & "    ,estado    = '" & l_estado & "' "
	l_sql = l_sql & "    ,last_updated_by = '" &session("loguinUser") & "'"
	l_sql = l_sql & "    ,last_update_date = GETDATE()" 
	l_sql = l_sql & " WHERE id = " & l_id

	l_cm.activeconnection = Cn
	l_cm.CommandText = l_sql
	cmExecute l_cm, l_sql, 0	
	
	' Inserto en la tabla historial	
	l_sql = "SELECT CONVERT(VARCHAR(8), fechahorainicio, 108) AS fechahorainicio, CONVERT(VARCHAR(10), fechahorainicio, 101) AS DateOnly  "
	l_sql = l_sql & " FROM calendarios "
	l_sql = l_sql & " WHERE id = " & l_id
	rsOpen l_rs, cn, l_sql, 0
	if not l_rs.eof then
		l_fecha = l_rs("DateOnly")
		l_hora = l_rs("fechahorainicio")
	end if
	l_rs.Close	
	
	l_evento = "Bloquear un Calendario"
	l_responsable = "Responsable"

	l_sql = "INSERT INTO historial_turnos "
	l_sql = l_sql & " (idcalendario, fechahorainicio, idrecursoreservable, idclientepaciente, evento, responsable, comentario ,  empnro, created_by,creation_date,last_updated_by,last_update_date)"
	l_sql = l_sql & " VALUES (" & l_id & "," & cambiaformato (l_fecha,l_hora ) & "," & l_idrecursoreservable & ",0,'" & l_evento &  "','" & l_responsable &  "','" & l_motivo &  "','" & session("empnro") & "','" & session("loguinUser")&"',GETDATE(),'"&session("loguinUser")&"',GETDATE())"
	l_cm.activeconnection = Cn
	l_cm.CommandText = l_sql
	cmExecute l_cm, l_sql, 0

	
else

' Response.write "<script>alert('Operación " &l_fechadesde&" Realizada.');</script>"
' Response.write "<script>alert('Operación " &l_horadesde&" Realizada.');</script>"

' Response.write "<script>alert('Operación " &l_fechahasta&" Realizada.');</script>"
' Response.write "<script>alert('Operación " &l_horahasta&" Realizada.');</script>"

	l_sql = "UPDATE calendarios "
	l_sql = l_sql & " SET motivo    = '" & l_motivo & "'"
	l_sql = l_sql & "    ,estado    = '" & l_estado & "' "
	l_sql = l_sql & "    ,last_updated_by = '" &session("loguinUser") & "'"
	l_sql = l_sql & "    ,last_update_date = GETDATE()" 	
	l_sql = l_sql & " WHERE CONVERT(VARCHAR(10), calendarios.fechahorainicio, 101) >=" & cambiafecha (l_fechadesde,true,1 )
	l_sql = l_sql & " AND CONVERT(VARCHAR(10), calendarios.fechahorainicio, 101) <=" & cambiafecha (l_fechahasta,true,1 )	  
	l_sql = l_sql & " AND CONVERT(VARCHAR(5), fechahorainicio, 108) >= '" & l_horadesde & "'"   
	l_sql = l_sql & " AND CONVERT(VARCHAR(5), fechahorainicio, 108) <= '" & l_horahasta & "'"
  
	if l_tipo = "B" then ' Bloquear
		l_sql = l_sql & " AND estado='ACTIVO'"
	else
		l_sql = l_sql & " AND estado='ANULADO'"
	end if
	l_sql = l_sql & " AND idrecursoreservable=" & l_idrecursoreservable
	l_sql = l_sql & " and calendarios.empnro = " & Session("empnro") 

	l_cm.activeconnection = Cn
	l_cm.CommandText = l_sql
	cmExecute l_cm, l_sql, 0		

end if





Set l_cm = Nothing

Response.write "<script>alert('Operación Realizada.');window.parent.opener.ifrm.location.reload();window.parent.close();</script>"

%>


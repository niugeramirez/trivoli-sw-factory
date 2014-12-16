<% Option Explicit %>
<!--#include virtual="/ticket/shared/inc/sec.inc"-->
<!--#include virtual="/ticket/shared/inc/const.inc"-->
<!--#include virtual="/ticket/shared/db/conn_db.inc"-->
<%
'Archivo: param_calidad_con_03.asp
'Descripción: Abm de parametros de calidad
'Autor : Lisandro Moro
'Fecha: 28/02/2005

Dim l_tipo
Dim l_rs
Dim l_cm
Dim l_sql

Dim l_camnro
Dim l_muecam
Dim l_etidescamsec
Dim l_etidescamhum
Dim l_etisal
Dim l_eticie
Dim l_paleo
Dim l_fumiga
Dim l_secado
Dim l_meralkilo
Dim l_tramuecam
Dim l_traimpeti

on error goto 0 
l_camnro = request.form("camnro")
l_muecam = request.form("muecam")
l_etidescamsec = request.form("etidescamsec")
l_etidescamhum = request.form("etidescamhum")
l_etisal = request.form("etisal")
l_eticie = request.form("eticie")
l_paleo = request.form("paleo")
l_fumiga = request.form("fumiga")
l_secado = request.form("secado")
l_meralkilo = request.form("meralkilo")
l_tramuecam = request.form("tramuecam")
l_traimpeti = request.form("traimpeti")

if l_etidescamsec = "" then
	l_etidescamsec = "null"
end if
if l_etidescamhum = "" then
	l_etidescamhum = "null"
end if
if l_etisal = "" then
	l_etisal = "null"
end if
if l_eticie = "" then
	l_eticie = "null"
end if


if l_muecam <> "" then
	l_muecam = -1
else
	l_muecam = 0
end if
if l_paleo <> "" then
	l_paleo = -1
else
	l_paleo = 0
end if
if l_fumiga <> "" then
	l_fumiga = -1
else
	l_fumiga = 0
end if
if l_secado <> "" then
	l_secado = -1
else
	l_secado = 0
end if
if l_meralkilo <> "" then
	l_meralkilo = -1
else
	l_meralkilo = 0
end if
if l_tramuecam <> "" then
	l_tramuecam = -1
else
	l_tramuecam = 0
end if
if l_traimpeti <> "" then
	l_traimpeti = -1
else
	l_traimpeti = 0
end if

set l_cm = Server.CreateObject("ADODB.Command")
l_cm.activeconnection = Cn
l_cm.CommandText = l_sql
l_sql = " UPDATE tkt_config "
l_sql = l_sql & " SET camnro = " & l_camnro 
l_sql = l_sql & " ,muecam = " & l_muecam
l_sql = l_sql & " ,etidescamsec = " & l_etidescamsec
l_sql = l_sql & " ,etidescamhum = " & l_etidescamhum
l_sql = l_sql & " ,etisal = " & l_etisal
l_sql = l_sql & " ,eticie = " & l_eticie
l_sql = l_sql & " ,paleo = " & l_paleo
l_sql = l_sql & " ,fumiga = " & l_fumiga
l_sql = l_sql & " ,secado = " & l_secado
l_sql = l_sql & " ,meralkilo = " & l_meralkilo
l_sql = l_sql & " ,tramuecam = " & l_tramuecam
l_sql = l_sql & " ,traimpeti = " & l_traimpeti
cmExecute l_cm, l_sql, 0


Set l_cm = Nothing

Response.write "<script>alert('Operación Realizada.');window.parent.parent.close();</script>"
%>

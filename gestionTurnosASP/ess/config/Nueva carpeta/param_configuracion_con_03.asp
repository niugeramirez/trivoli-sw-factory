<% Option Explicit %>
<!--#include virtual="/ticket/shared/inc/sec.inc"-->
<!--#include virtual="/ticket/shared/inc/const.inc"-->
<!--#include virtual="/ticket/shared/db/conn_db.inc"-->
<%
'Archivo: param_configuracion_con_03.asp
'Descripción: Abm de parametros de configuracion
'Autor : Lisandro Moro
'Fecha: 01/03/2005
' Modificado: Raul CHinestra - 15/06/2006 - Se agregó la cantidad de dias que se desea mostrar el transito

Dim l_tipo
Dim l_rs
Dim l_cm
Dim l_sql

Dim l_configura
Dim l_productiv
Dim l_habilita
Dim l_casrem
Dim l_clapes
Dim l_merxvag
Dim l_difpesxven
Dim l_operati
Dim l_proxlin
Dim l_bloconcum
Dim l_nroins
Dim l_netdpcam
Dim l_netdpvag
Dim l_brudpcam
Dim l_brudpvag
Dim l_summez
Dim l_brumaxcam
Dim l_ticpre
Dim l_rempre
Dim l_cantic
Dim l_canrem
Dim l_etiosob
Dim l_mostra
Dim l_cirpla

'dim algo
'for each algo in request.form
'	Response.Write "l_" & algo & "=" & request.form(algo)&"<br>"
'next
'on error goto 0 

'Response.End

l_configura = request.form("configura")
l_merxvag = request.form("merxvag")
l_netdpcam = request.form("netdpcam")
l_brudpcam = request.form("brudpcam")
l_productiv = request.form("productiv")
l_habilita = request.form("habilita")
l_operati = request.form("operati")
l_rempre = request.form("rempre")
l_cantic = request.form("cantic")
l_casrem = request.form("casrem")
l_ticpre = request.form("ticpre")
l_clapes = request.form("clapes")
l_bloconcum = request.form("bloconcum")
l_difpesxven = request.form("difpesxven")
l_proxlin = request.form("proxlin")
l_netdpvag = request.form("netdpvag")
l_brumaxcam = request.form("brumaxcam")
l_nroins = request.form("nroins")
l_summez = request.form("summez")
l_brudpvag = request.form("brudpvag")
l_canrem = request.form("canrem")
l_etiosob = request.form("etiosob")
l_mostra = request.form("mostra")
l_cirpla = request.form("cirpla")

'if l_configura = "" then l_configura = null
if l_merxvag <> "" then l_merxvag = -1 else l_merxvag = 0
if l_netdpcam = "" then l_netdpcam = "null"
if l_brudpcam = "" then l_brudpcam = "null"
if l_productiv <> "" then l_productiv = -1 else l_productiv = 0
if l_habilita <> "" then l_habilita = -1 else l_habilita = 0
if l_operati <> "" then l_operati = -1 else l_operati = 0
if l_rempre <> "" then l_rempre = -1 else l_rempre = 0
if l_cantic = "" then l_cantic = "null"
'if l_casrem 
if l_ticpre <> "" then l_ticpre = -1 else l_ticpre = 0
if l_clapes <> "" then l_clapes = -1 else l_clapes = 0
if l_bloconcum <> "" then l_bloconcum = -1 else l_bloconcum = 0
if l_difpesxven <> "" then l_difpesxven = -1 else l_difpesxven = 0
if l_proxlin <> "" then l_proxlin = -1 else l_proxlin = 0
if l_netdpvag = "" then l_netdpvag = "null"
if l_brumaxcam = "" then l_brumaxcam = "null"
if l_nroins = "" then l_nroins = "null"
if l_summez = "" then l_summez = "null"
if l_brudpvag = "" then l_brudpvag = "null"
if l_canrem = "" then l_canrem = "null"
'l_etiosob
'l_cirpla 

set l_cm = Server.CreateObject("ADODB.Command")
l_cm.activeconnection = Cn
l_cm.CommandText = l_sql
l_sql = " UPDATE tkt_config "
l_sql = l_sql & " set configura = '" & l_configura & "'"
l_sql = l_sql & " ,merxvag = " & l_merxvag
l_sql = l_sql & " ,netdpcam = " & l_netdpcam
l_sql = l_sql & " ,brudpcam = " & l_brudpcam
l_sql = l_sql & " ,productiv = " & l_productiv
l_sql = l_sql & " ,habilita = " & l_habilita
l_sql = l_sql & " ,operati = " & l_operati
l_sql = l_sql & " ,rempre = " & l_rempre
l_sql = l_sql & " ,cantic = " & l_cantic
l_sql = l_sql & " ,casrem = '" & l_casrem & "'"
l_sql = l_sql & " ,ticpre = " & l_ticpre
l_sql = l_sql & " ,clapes = " & l_clapes
l_sql = l_sql & " ,bloconcum = " & l_bloconcum
l_sql = l_sql & " ,difpesxven = " & l_difpesxven
l_sql = l_sql & " ,proxlin = " & l_proxlin
l_sql = l_sql & " ,netdpvag = " & l_netdpvag
l_sql = l_sql & " ,brumaxcam = " & l_brumaxcam 
l_sql = l_sql & " ,nroins = " & l_nroins 
l_sql = l_sql & " ,summez = " & l_summez
l_sql = l_sql & " ,brudpvag = " & l_brudpvag
l_sql = l_sql & " ,canrem = " & l_canrem
l_sql = l_sql & " ,etiosob = '" & l_etiosob & "'"
l_sql = l_sql & " ,mostra = " & l_mostra
l_sql = l_sql & " ,cirpla = '" & ucase(l_cirpla) & "'"
cmExecute l_cm, l_sql, 0

Set l_cm = Nothing

Response.write "<script>alert('Operación Realizada.');window.parent.parent.close();</script>"
%>

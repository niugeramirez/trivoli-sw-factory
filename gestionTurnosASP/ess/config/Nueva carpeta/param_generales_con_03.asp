<% Option Explicit %>
<!--#include virtual="/ticket/shared/inc/sec.inc"-->
<!--#include virtual="/ticket/shared/inc/const.inc"-->
<!--#include virtual="/ticket/shared/db/conn_db.inc"-->
<%

on error goto 0

'Archivo: param_generales_con_03.asp
'Descripción: Abm de parametros de generales
'Autor : Lisandro Moro
'Fecha: 28/02/2005

Dim l_tipo
Dim l_rs
Dim l_cm
Dim l_sql

Dim l_empnro 
Dim l_lugnro 
Dim l_entnro 
Dim l_vennro
Dim l_recnro 
'Dim l_humcam 
'Dim l_humvag 
'Dim l_humdir 
Dim l_dismerhum 
Dim l_traemp 
Dim l_balcodcon 
Dim l_pesdespro 
Dim l_txtpla 
Dim l_txtcup 
Dim l_promov 
Dim l_propla 
Dim l_protra 
Dim l_desnro 
Dim l_mosrec 
Dim l_mosnrotap 
Dim l_mosmez 
Dim l_cupplaya 
Dim l_txtcos
Dim l_meralkilo
Dim l_carporfersuc
Dim l_carporfernum

'dim algo
'for each algo in request.form
'	Response.Write algo &"=" & request.form(algo) & "<br>"
'next
'Response.End

on error goto 0 

l_empnro = request.form("empnro")
l_lugnro = request.form("lugnro")
l_entnro = request.form("entnro")
l_vennro = request.form("vennro")
l_recnro = request.form("recnro")
'l_humcam = request.form("humcam")
'l_humvag = request.form("humvag")
'l_humdir = request.form("humdir")
l_dismerhum = request.form("dismerhum")
l_traemp = request.form("traemp")
l_balcodcon = request.form("balcodcon")
l_pesdespro = request.form("pesdespro")
l_txtpla = request.form("txtpla")
l_txtcup = request.form("txtcup")
l_promov = request.form("promov")
l_propla = request.form("propla")
l_protra = request.form("protra")
l_desnro = request.form("desnro")
l_mosrec = request.form("mosrec")
l_mosnrotap = request.form("mosnrotap")
l_mosmez = request.form("mosmez")
l_cupplaya = request.form("cuppla")
l_txtcos = request.form("txtcos")
'l_meralkilo = request.form("meralkilo")
l_carporfersuc = request.form("carporfersuc")
l_carporfernum = request.form("carporfernum")

'l_empnro
'l_lugnro
'l_entnro
'l_recnro
'if l_humcam <> "" then l_humcam = -1 else l_humcam = 0
'if l_humvag <> "" then l_humvag = -1 else l_humvag = 0
'if l_humdir = "" then l_humdir = null
if l_dismerhum <> "" then l_dismerhum = -1 else l_dismerhum = 0
if l_traemp <> "" then l_traemp = -1 else l_traemp = 0
if l_balcodcon = "" then l_balcodcon = "null"
if l_pesdespro = "" then l_pesdespro  = null
if l_txtpla = "" then l_txtpla = null
if l_txtcup = "" then l_txtcup = null
if l_promov = "" then l_promov = "null"
if l_propla = "" then l_propla = "null"
if l_protra = "" then l_protra = "null"
if l_txtcos = "" then l_txtcos = ""
'l_desnro
if l_mosrec <> "" then l_mosrec = -1 else l_mosrec = 0
if l_mosnrotap <> "" then l_mosnrotap = -1 else l_mosnrotap = 0
if l_mosmez <> "" then l_mosmez = -1 else l_mosmez = 0
if l_cupplaya <> "" then l_cupplaya = -1 else l_cupplaya = 0
'if l_meralkilo <> "" then l_meralkilo = -1 else l_meralkilo = 0

set l_cm = Server.CreateObject("ADODB.Command")
l_cm.activeconnection = Cn
l_cm.CommandText = l_sql
l_sql = " UPDATE tkt_config "
l_sql = l_sql & " set empnro = " & l_empnro 
l_sql = l_sql & " ,lugnro = " & l_lugnro 
l_sql = l_sql & " ,entnro = " & l_entnro 
l_sql = l_sql & " ,vennro = " & l_vennro 
l_sql = l_sql & " ,recnro = " & l_recnro 
'l_sql = l_sql & " ,humcam = " & l_humcam  
'l_sql = l_sql & " ,humvag = " & l_humvag 
'l_sql = l_sql & " ,humdir = '" & l_humdir & "'"
l_sql = l_sql & " ,dismerhum = " & l_dismerhum 
l_sql = l_sql & " ,traemp = " & l_traemp 
l_sql = l_sql & " ,balcodcon = " & l_balcodcon
l_sql = l_sql & " ,pesdespro = '" & l_pesdespro & "'"
l_sql = l_sql & " ,txtpla = '" & l_txtpla & "'"
l_sql = l_sql & " ,txtcup = '" & l_txtcup  & "'"
l_sql = l_sql & " ,promov = " &  l_promov 
l_sql = l_sql & " ,propla = " &  l_propla 
l_sql = l_sql & " ,protra = " &  l_protra 
l_sql = l_sql & " ,desnro = " & l_desnro 
l_sql = l_sql & " ,mosrec = " & l_mosrec 
l_sql = l_sql & " ,mosnrotap = " & l_mosnrotap
l_sql = l_sql & " ,mosmez = " & l_mosmez
l_sql = l_sql & " ,cupplaya = " & l_cupplaya
l_sql = l_sql & " ,txtcos = '" & l_txtcos & "'"
'l_sql = l_sql & " ,meralkilo = " & l_meralkilo
l_sql = l_sql & " ,carporfersuc = " & l_carporfersuc
l_sql = l_sql & " ,carporfernum = " & l_carporfernum
cmExecute l_cm, l_sql, 0

Set l_cm = Nothing

Response.write "<script>alert('Operación Realizada.');window.parent.parent.close();</script>"
%>

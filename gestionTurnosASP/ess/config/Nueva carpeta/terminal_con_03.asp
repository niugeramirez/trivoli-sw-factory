<% Option Explicit %>
<!--#include virtual="/ticket/shared/inc/sec.inc"-->
<!--#include virtual="/ticket/shared/inc/const.inc"-->
<!--#include virtual="/ticket/shared/db/conn_db.inc"-->
<% 
'Archivo: terminal_con_03.asp
'Descripción: ABM de terminales
'Autor : Gustavo Manfrin
'Fecha: 15/04/2005
'on error goto 0
Dim l_tipo
Dim l_cm
Dim l_sql

Dim l_ternro
Dim l_terdes
Dim l_tercod
Dim l_tercov
Dim l_tersect
Dim l_terbalf1
Dim	l_balanza1
Dim l_tercomf1
Dim l_tercomfcf1
Dim l_tervvcf1
Dim l_terbalf2
Dim	l_balanza2
Dim l_tercomf2
Dim l_tercomfcf2
Dim l_tervvcf2
Dim l_terimptick
Dim l_terimpcpor
Dim l_terimpremi
Dim l_terimpetiq
Dim l_terresmacp
Dim l_planro
Dim l_terizq

l_tipo 		= request.querystring("tipo")
l_ternro 	= request.Form("ternro")
l_terdes = request.Form("terdes")
l_tercod = request.Form("tercod")
l_tercov = request.Form("tercov")
l_tersect = request.Form("tersec")
l_terbalf1 = request.Form("terbal1")
l_tercomf1 = request.Form("tercom1")
l_tercomfcf1 = request.Form("tercomfc1")
l_tervvcf1 = request.Form("tervvc1")
l_terbalf2 = request.Form("terbal2")
l_tercomf2 = request.Form("tercom2")
l_tercomfcf2 = request.Form("tercomfc2")
l_tervvcf2 = request.Form("tervvc2")
l_terimptick = request.Form("teritkt")
l_terimpcpor = request.Form("tericpor")
l_terimpremi = request.Form("teriremi")
l_terimpetiq = request.Form("terieti")
l_terresmacp = request.Form("terrcpor")
l_planro = request.Form("planro")
l_terizq = request.Form("izq")

if l_terbalf1 = "" then
	l_terbalf1 = "null"
else
	l_tervvcf1 = -1
end if 

if l_tercomf1 = "" then
	l_tercomf1 = "null"
end if 

if l_tercomfcf1 = "" then
	l_tercomfcf1 = "null"
end if 

if l_tervvcf1 = "" then
	l_tervvcf1 = "null"
else
	l_tervvcf1 = -1
end if 

if l_terbalf2 = "" then
	l_terbalf2 = "null"
else
	l_tervvcf2 = -1
end if 

if l_tercomf2 = "" then
	l_tercomf2 = "null"
end if 

if l_tercomfcf2 = "" then
	l_tercomfcf2 = "null"
end if '''

if l_tervvcf2 = "" then
	l_tervvcf2 = "null"
else
	l_tervvcf2 = -1
end if 

if l_terresmacp = "" then
	l_terresmacp = "1"
end if 


	set l_cm = Server.CreateObject("ADODB.Command")
	if l_tipo = "A" then 
		l_sql = "INSERT INTO tkt_terminal"
		l_sql = l_sql & " (terdesc, tercod, tercov, tersect, terbalf1, tercomf1, tercomfcf1, tervvcf1,"
		l_sql = l_sql & " terbalf2, tercomf2, tercomfcf2, tervvcf2,terimptick, terimpcpor, terimpremi, terimpetiq,terresmacp,planro) " ', terizq)"
		l_sql = l_sql & " VALUES ('" & l_terdes 
		l_sql = l_sql & "','" & l_tercod & "','" & l_tercov & "','" & l_tersect & "'," & l_terbalf1 
    	l_sql = l_sql & "," & l_tercomf1 & "," & l_tercomfcf1 & "," & l_tervvcf1 & "," & l_terbalf2 
    	l_sql = l_sql & "," & l_tercomf2 & "," & l_tercomfcf2 & "," & l_tervvcf2 & ",'" & l_terimptick 
    	l_sql = l_sql & "','" & l_terimpcpor & "','" & l_terimpremi & "','" & l_terimpetiq & "'," & l_terresmacp &"," & l_planro & ")"
'		l_sql = l_sql & ",'" & l_terizq & "')"
	else
		l_sql = "UPDATE tkt_terminal"
		l_sql = l_sql & " SET terdesc = '" & l_terdes & "'"
		l_sql = l_sql & ", tercod = '" & l_tercod & "'"
		l_sql = l_sql & ", tercov = '" & l_tercov & "'"
		l_sql = l_sql & ", tersect = '" & l_tersect & "'"
		l_sql = l_sql & ", terbalf1 = " & l_terbalf1 
		l_sql = l_sql & ", tercomf1 = " & l_tercomf1 		
		l_sql = l_sql & ", tercomfcf1 = " & l_tercomfcf1 		
		l_sql = l_sql & ", tervvcf1 = " & l_tervvcf1 		
		l_sql = l_sql & ", terbalf2 = " & l_terbalf2 
		l_sql = l_sql & ", tercomf2 = " & l_tercomf2 		
		l_sql = l_sql & ", tercomfcf2 = " & l_tercomfcf2 		
		l_sql = l_sql & ", tervvcf2 = " & l_tervvcf2 		
		l_sql = l_sql & ", terimptick = '" & l_terimptick & "'"		
		l_sql = l_sql & ", terimpcpor = '" & l_terimpcpor & "'"				
		l_sql = l_sql & ", terimpremi = '" & l_terimpremi & "'"				
		l_sql = l_sql & ", terimpetiq = '" & l_terimpetiq & "'"				
		l_sql = l_sql & ", terresmacp = '" & l_terresmacp & "'"				
		l_sql = l_sql & ", planro = " & l_planro				
'		l_sql = l_sql & ", terizq = '" & l_terizq & "'"
  	    l_sql = l_sql & " WHERE ternro = " & l_ternro
	end if
	'response.write l_sql & "<br>"
	l_cm.activeconnection = Cn
	l_cm.CommandText = l_sql
	cmExecute l_cm, l_sql, 0
	Set l_cm = Nothing

	Response.write "<script>alert('Operación Realizada.');window.parent.opener.ifrm.location.reload();window.parent.close();</script>"
%>


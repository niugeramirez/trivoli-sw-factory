<% Option Explicit %>
<!--#include virtual="/serviciolocal/shared/inc/sec.inc"-->
<!--#include virtual="/serviciolocal/shared/inc/const.inc"-->
<!--#include virtual="/serviciolocal/shared/db/conn_db.inc"-->
<!--#include virtual="/serviciolocal/shared/inc/fecha.inc"-->
<% 
'Archivo: depositos_con_03.asp
'Descripción: ABM de Depósitos
'Autor : Raul Chinestra	
'Fecha: 11/02/2005

'on error resume next

Dim l_tipo
Dim l_cm
Dim l_sql
Dim l_rs2

Dim l_empremsuc
Dim l_empremnro
Dim l_emp1suc
Dim l_emp1nro
Dim l_emp2suc
Dim l_emp2nro
Dim l_emp3suc
Dim l_emp3nro

Dim l_empnro


Dim l_empcau1nro
Dim l_empcau2nro
Dim l_empcau3nro
Dim l_empfecvto1
Dim l_empfecvto2
Dim l_empfecvto3

Dim l_empcai
Dim l_empfecini
Dim l_empfecven
Dim l_empticpre
Dim l_dirloc
Dim l_locloc
Dim l_locloccod

Dim l_empesp

l_empnro 	= request.Form("empnro")

l_empcai	= request.Form("empcai")
l_empfecini = request.Form("empfecini")
l_empfecven = request.Form("empfecven")
l_empticpre = request.Form("empticpre")

l_empremsuc = request.Form("empremsuc")
l_empremnro = request.Form("empremnro")
l_emp1suc	= request.Form("empcarpor1suc")
l_emp1nro	= request.Form("empcarpor1nro")
l_emp2suc	= request.Form("empcarpor2suc")
l_emp2nro	= request.Form("empcarpor2nro")
l_emp3suc	= request.Form("empcarpor3suc")
l_emp3nro	= request.Form("empcarpor3nro")

l_empcau1nro = request.Form("empcau1nro")
l_empfecvto1 = request.Form("empfecvto1")
l_empcau2nro = request.Form("empcau2nro")
l_empfecvto2 = request.Form("empfecvto2")
l_empcau3nro = request.Form("empcau3nro")
l_empfecvto3 = request.Form("empfecvto3")

l_empesp = request.Form("empesp")

l_dirloc = request.Form("dirloc")
l_locloccod = request.Form("locloccod")

set l_cm = Server.CreateObject("ADODB.Command")

if l_empticpre = "" then
	l_empticpre = "null"
end if

if l_emp1suc = "" then
	l_emp1suc = "null"
end if 

if l_empremsuc = "" then
	l_empremsuc = "null"
end if 

if l_empremnro = "" then
 	l_empremnro = "null"
end if 


if l_emp1nro = "" then
	l_emp1nro = "null"
end if 

if l_emp2suc = "" then
	l_emp2suc = "null"
end if 

if l_emp2nro = "" then
	l_emp2nro = "null"
end if 

if l_emp3suc = "" then
	l_emp3suc = "null"
end if 

if l_emp3nro = "" then
	l_emp3nro = "null"
end if 

if l_locloccod = "" then
	l_locloc = "null"
else
	Set l_rs2 = Server.CreateObject("ADODB.RecordSet")

	l_sql = "SELECT locnro FROM tkt_localidad "
	l_sql  = l_sql  & " WHERE loccod = '" & l_locloccod  & "'"
	
	rsOpen l_rs2, cn, l_sql, 0 
	
	if not l_rs2.eof then
		l_locloc = l_rs2("locnro")  
	else
		l_locloc = "null"
	end if
	l_rs2.Close
	Set l_rs2 = Nothing
end if 

'response.write l_locloc & "<br>"

l_sql = "UPDATE tkt_empresa "
l_sql = l_sql & " SET empcarpor1suc = " & l_emp1suc
l_sql = l_sql & ", empcai = '" & l_empcai & "'"
if l_empfecini = "" or isnull(l_empfecini) then 
	l_sql = l_sql & ", empfecini = null" 
else
	l_sql = l_sql & ", empfecini = " & cambiafecha(l_empfecini,"YMD",true)
end if
if l_empfecven = "" or isnull(l_empfecven) then 
	l_sql = l_sql & ", empfecven = null" 
else
	l_sql = l_sql & ", empfecven = " & cambiafecha(l_empfecven,"YMD",true)
end if
l_sql = l_sql & ", empticpre = " & l_empticpre 
l_sql = l_sql & ", empremsuc = " & l_empremsuc
l_sql = l_sql & ", empremnro = " & l_empremnro 
l_sql = l_sql & ", empcarpor1nro = " & l_emp1nro 
l_sql = l_sql & ", empcarpor2suc = " & l_emp2suc 
l_sql = l_sql & ", empcarpor2nro = " & l_emp2nro 
l_sql = l_sql & ", empcarpor3suc = " & l_emp3suc 
l_sql = l_sql & ", empcarpor3nro = " & l_emp3nro 
l_sql = l_sql & ", empcau1nro = '" & l_empcau1nro & "'"
if l_empfecvto1 = "" or isnull(l_empfecvto1) then 
	l_sql = l_sql & ", empfecvto1 = null" 
else
	l_sql = l_sql & ", empfecvto1 = " & cambiafecha(l_empfecvto1,"YMD",true)
end if
l_sql = l_sql & ", empcau2nro = '" & l_empcau2nro & "'"
if l_empfecvto2 = "" or isnull(l_empfecvto2) then 
	l_sql = l_sql & ", empfecvto2 = null" 
else
	l_sql = l_sql & ", empfecvto2 = " & cambiafecha(l_empfecvto2,"YMD",true)
end if
l_sql = l_sql & ", empcau3nro = '" & l_empcau3nro & "'"
if l_empfecvto3 = "" or isnull(l_empfecvto3) then 
	l_sql = l_sql & ", empfecvto3 = null" 
else
	l_sql = l_sql & ", empfecvto3 = " & cambiafecha(l_empfecvto3,"YMD",true)
end if
if l_empesp = "on" then 
   l_sql = l_sql & ", empesp = -1"
else 
   l_sql = l_sql & ", empesp = 0"
end if 	
l_sql = l_sql & ", dirloc = '" & l_dirloc & "'"
l_sql = l_sql & ", locloc = " & l_locloc 
l_sql = l_sql & " WHERE empnro = " & l_empnro

'response.write l_sql & "<br>"
l_cm.activeconnection = Cn
l_cm.CommandText = l_sql
cmExecute l_cm, l_sql, 0
Set l_cm = Nothing

Response.write "<script>alert('Operación Realizada.');window.parent.opener.ifrm.location.reload();window.parent.close();</script>"
%>


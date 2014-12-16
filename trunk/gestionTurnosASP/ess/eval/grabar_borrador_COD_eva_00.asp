<% Option Explicit %>
<!--#include virtual="/serviciolocal/shared/inc/sec.inc"-->
<!--#include virtual="/serviciolocal/shared/inc/const.inc"-->
<!--#include virtual="/serviciolocal/shared/db/conn_db.inc"-->
<!--#include virtual="/serviciolocal/shared/inc/fecha.inc"-->
<!--#include virtual="/serviciolocal/shared/inc/sqls.inc"-->
<!--#include virtual="/serviciolocal/shared/inc/const_eva.inc"-->
<% 
'=====================================================================================
'Archivo  : grabar_borrador_COD_eva_00.asp
'Objetivo : grabar de objetivos borrador
'Fecha	  : 01-02-2005
'Autor	  : CCRossi
'=====================================================================================

' variables
' parametros de entrada ----------------------------------------
  Dim l_evldrnro
  Dim l_evaobjborrnro
  Dim l_evaobjborrdext
  Dim l_evaobjborrfmed
  Dim l_evaobjborrpond
  Dim l_evatipobjnro
  
' variables de base de datos ------------------------------------
  Dim l_cm
  Dim l_sql
  Dim l_rs
  Dim l_rs1
  Dim l_tipo  
  
' locales
  Dim l_evacabnro 
  Dim l_evatevnro
  Dim l_evaluador

  Dim l_suma
  Dim l_cartel
    
' parametros de entrada
  l_evaobjborrdext  = left(trim(request.querystring("evaobjborrdext")),200)
  l_evaobjborrfmed	= left(trim(request.querystring("evaobjborrfmed")),200)
  l_evldrnro		= request.querystring("evldrnro")
  l_tipo			= request.querystring("tipo")
  l_evaobjborrnro	= request.querystring("evaobjborrnro")
  l_evaobjborrpond	= request.querystring("evaobjborrpond")
  l_evatipobjnro	= request.querystring("evatipobjnro")

if trim(l_evaobjborrfmed)="" then
l_evaobjborrfmed=" "
else
l_evaobjborrfmed= trim(l_evaobjborrfmed)
end if
'BODY ----------------------------------------------------------
select case l_tipo
Case "A":
		' Primero sumar lo que hay para ver si excede el 100%
		Set l_rs = Server.CreateObject("ADODB.RecordSet")
		l_sql = "SELECT evaobjborrpond "
		l_sql = l_sql & " FROM evaobjborr "
		l_sql = l_sql & " INNER JOIN evaluaobjborr ON evaluaobjborr.evaobjborrnro = evaobjborr.evaobjborrnro"
		l_sql = l_sql & " LEFT  JOIN evatipoobj ON evatipoobj.evatipobjnro = evaobjborr.evatipobjnro"
		l_sql = l_sql & " WHERE evatipoobj.evatipobjnro = " & l_evatipobjnro
		l_sql = l_sql & " AND   evaluaobjborr.evldrnro = " & l_evldrnro
		'Response.Write l_sql
		rsOpen l_rs, cn, l_sql, 0 
		l_suma=0
		do while not l_rs.eof
			if trim(l_rs("evaobjborrpond"))<>"" then
			l_suma= l_suma + cdbl(l_rs("evaobjborrpond"))
			end if
			l_rs.MoveNext
		loop
		l_rs.close
		set l_rs=nothing

		if (l_suma + l_evaobjborrpond) > 100 then
			l_evaobjborrpond = 0
			l_cartel = " Modifique la Ponderación del nuevo Compromiso.\n La sumatoria no puede superar 100%."
		end if
		
		l_sql= "insert into evaobjborr (evaobjborrdext,evaobjborrfmed,evatipobjnro, evaobjborrpond) "
		l_sql = l_sql & " values ('" & trim(l_evaobjborrdext) & "','" & l_evaobjborrfmed & "'," & l_evatipobjnro & "," & cint(l_evaobjborrpond) &")"
case "M":
		l_sql = "UPDATE evaobjborr SET "
		l_sql = l_sql & " evaobjborrdext   = '" & trim(l_evaobjborrdext) & "',"
		l_sql = l_sql & " evaobjborrfmed = '" & l_evaobjborrfmed & "', "
		l_sql = l_sql & " evaobjborrpond = " & cint(l_evaobjborrpond) 
		l_sql = l_sql & " WHERE evaobjborr.evaobjborrnro = "  & l_evaobjborrnro
case "B":
		l_sql = "DELETE from evaluaobjborr where evaobjborrnro = "  & l_evaobjborrnro
		set l_cm = Server.CreateObject("ADODB.Command")  
		l_cm.activeconnection = Cn
		l_cm.CommandText = l_sql
		cmExecute l_cm, l_sql, 0
	
		l_sql = "DELETE from evaobjborr where evaobjborr.evaobjborrnro = "  & l_evaobjborrnro
end select

set l_cm = Server.CreateObject("ADODB.Command")  
l_cm.activeconnection = Cn
l_cm.CommandText = l_sql
cmExecute l_cm, l_sql, 0

'response.write l_sql

if trim(l_evaobjborrnro)="" then
	Set l_rs = Server.CreateObject("ADODB.RecordSet")
	l_sql = fsql_seqvalue("evaobjborrnro","evaobjborr")
	rsOpen l_rs, cn, l_sql, 0
	if not l_rs.eof then
		l_evaobjborrnro=l_rs("evaobjborrnro")
	end if	
	l_rs.Close
	Set l_rs = Nothing
end if

'Crear registro de la relacion ================================================
select case l_tipo
Case "A":
	Set l_rs1 = Server.CreateObject("ADODB.RecordSet")
	l_sql = "SELECT * FROM evaluaobjborr "
	l_sql = l_sql & " WHERE evaobjborrnro = " & l_evaobjborrnro
	l_sql = l_sql & " AND   evldrnro  = " & l_evldrnro
	rsOpen l_rs1, cn, l_sql, 0
	if  l_rs1.eof then
		l_sql= "insert into evaluaobjborr (evldrnro,evaobjborrnro) "
		l_sql = l_sql & " values (" & l_evldrnro & "," & l_evaobjborrnro &")"
		set l_cm = Server.CreateObject("ADODB.Command")  
		l_cm.activeconnection = Cn
		l_cm.CommandText = l_sql
		cmExecute l_cm, l_sql, 0
	end if	
	l_rs1.Close
	set l_rs1=nothing

end select

if trim(l_cartel)<>"" then
	response.write "<script> alert('"&l_cartel&"'); </script>"
end if

cn.close
set cn=nothing

response.write "<script> parent.location.reload(); </script>"

		
%>

<% Option Explicit %>
<!--#include virtual="/serviciolocal/shared/inc/sec.inc"-->
<!--#include virtual="/serviciolocal/shared/inc/const.inc"-->
<!--#include virtual="/serviciolocal/shared/inc/const_eva.inc"-->
<!--#include virtual="/serviciolocal/shared/db/conn_db.inc"-->
<!--#include virtual="/serviciolocal/shared/inc/fecha.inc"-->
<% 
'Modificado	: 18-03-2005 controlar grabacion de fechas planificacion, seguimiento para CODELCO

' variables
' parametros de entrada ----------------------------------------

  on error goto 0

  Dim l_evaevenro
  Dim l_tipo

' variables de base de datos ------------------------------------
  Dim l_cm
  Dim l_sql
  Dim l_rs

' variables que vienen del form
  dim l_evaevedesabr
  dim l_evaevedesext
  dim l_evaevefecha
  dim l_evaevefplan
  dim l_evaevefseg
  dim l_evaevefdesde
  dim l_evaevefhasta
  dim l_evatipnro
  dim l_evaperact
  dim l_evaperprox
  dim l_evatipevenro
  
' parametros de entrada
  l_tipo		  = Request.Form("tipo")
  l_evaevenro	  = Request.Form("evaevenro")	
  l_evaevedesabr  = request.Form("evaevedesabr")
  l_evaevedesext  = left(request.Form("evaevedesext"),200)
  l_evaevefecha   = request.Form("evaevefecha")
  l_evaevefdesde  = request.Form("evaevefdesde")
  l_evaevefhasta  = request.Form("evaevefhasta")
  l_evatipnro	  = request.Form("evatipnro")
  l_evaperact	  = request.Form("evaperact")
  l_evaperprox    = request.Form("evaperprox")
  l_evatipevenro  = request.Form("evatipevenro")
  l_evaevefplan   = request.Form("evaevefplan")
  l_evaevefseg    = request.Form("evaevefseg")
  

'BODY ----------------------------------------------------------

 l_evaevefecha   = cambiafecha(l_evaevefecha,"","")
 l_evaevefdesde  = cambiafecha(l_evaevefdesde,"","")
 l_evaevefhasta  = cambiafecha(l_evaevefhasta,"","")

if trim(l_evaevefplan) <> "" then
 l_evaevefplan   = cambiafecha(l_evaevefplan,"YMD",false)
end if
if trim(l_evaevefseg) <> "" then
 l_evaevefseg   = cambiafecha(l_evaevefseg,"YMD",false)
end if
  
' chequear que no haya  otro con la misma descripcion
  Set l_rs = Server.CreateObject("ADODB.RecordSet")
  l_sql = "SELECT evaevedesabr FROM  evaevento WHERE evaevedesabr = '" & l_evaevedesabr & "'"
  if l_tipo = "M" then  
		l_sql = l_sql & " AND evaevenro <> " & l_evaevenro
  end if		
  rsOpen l_rs, cn, l_sql, 0 
  if not l_rs.eof then
	Response.write "<script>alert('Ya existe otro Evento con la misma descripcion.');window.close();</script>"
  else
    set l_cm = Server.CreateObject("ADODB.Command")
	if l_tipo = "A" then 
		l_sql = "INSERT INTO evaevento (evaevedesabr, evaevedesext, evaevefecha, evaevefdesde, "
		if ccodelco=-1 then
		l_sql = l_sql & "  evaevefplan,evaevefseg,"
		end if
		l_sql = l_sql & "  evaevefhasta, evatipnro,evaperact,evaperprox, evatipevenro) "
		l_sql = l_sql & " VALUES ('" & l_evaevedesabr & "','" & l_evaevedesext   & "'," & l_evaevefecha  & "," & l_evaevefdesde & ","
		if ccodelco=-1 then
		l_sql = l_sql & l_evaevefplan  & "," & l_evaevefseg & ","
		end if
		l_sql = l_sql & l_evaevefhasta & ","  & l_evatipnro & "," 
		if len(trim(l_evaperact)) <> 0 then
			l_sql = l_sql & l_evaperact   & ","
		else		
			l_sql = l_sql & "null,"
		end if	   
		if len(trim(l_evaperprox)) <> 0 then
			l_sql = l_sql & l_evaperprox   & ","
		else		
			l_sql = l_sql & "null,"
		end if		
		l_sql = l_sql & l_evatipevenro & ")"
	else
		l_sql = "UPDATE  evaevento SET evaevedesabr ='" & l_evaevedesabr & "',evaevedesext = '" & l_evaevedesext & "', evaevefecha = " & l_evaevefecha  & " ," 
		if ccodelco=-1 then
		l_sql = l_sql & " evaevefplan = " & l_evaevefplan  & " ,evaevefseg = " & l_evaevefseg  & " ,"  
		end if
		l_sql = l_sql & " evaevefdesde=" & l_evaevefdesde & " ,evaevefhasta = " & l_evaevefhasta & " ,evatipnro = " & l_evatipnro & ", " 
		if len(trim(l_evaperact)) <> 0 then
			l_sql = l_sql & " evaperact   =  " & l_evaperact   & ", " 
		end if
		if len(trim(l_evaperprox)) <> 0 then
			l_sql = l_sql & " evaperprox   =  " & l_evaperprox   & ", " 
		end if
		l_sql = l_sql & " evatipevenro =  " & l_evatipevenro & " " & " WHERE evaevenro  = " & l_evaevenro
	end if
	'Response.Write(l_sql)
	l_cm.activeconnection = Cn
	l_cm.CommandText = l_sql
    cmExecute l_cm, l_sql, 0
    Response.write "<script>alert('Operación Realizada.');window.opener.ifrm.location.reload();window.close();</script>"
  end if
%>

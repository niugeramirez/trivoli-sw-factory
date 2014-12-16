<% Option Explicit %>
<!--#include virtual="/serviciolocal/shared/inc/sec.inc"-->
<!--#include virtual="/serviciolocal/shared/inc/const.inc"-->
<!--#include virtual="/serviciolocal/shared/inc/const_eva.inc"-->
<!--#include virtual="/serviciolocal/shared/db/conn_db.inc"-->
<!--#include virtual="/serviciolocal/shared/inc/fecha.inc"-->
<% 
' variables
' parametros de entrada ----------------------------------------
  Dim l_evldrnro
  Dim l_evareunion
  Dim l_evaacuerdo
  Dim l_evafecha
  Dim l_evaobser
  Dim l_evaetapa
  dim l_puntaje 
  dim l_puntajemanual
  
'locales
 dim	l_evacabnro 
 dim	l_evatevnro 
 dim	l_lista    
 
' variables de base de datos ------------------------------------
  Dim l_cm
  Dim l_sql
  Dim l_rs
    
' parametros de entrada
  l_evaobser  = left(trim(request.querystring("evaobser")),200)
  l_evldrnro = request.querystring("evldrnro")
  l_evafecha = request.querystring("evafecha")
  l_evareunion = request.querystring("evareunion")
  l_evaacuerdo = request.querystring("evaacuerdo")
  l_evaetapa   = request.querystring("evaetapa")
  
' si es etapa 3 tambien viene la notafinal modificada
 l_puntaje	 = request.querystring("notafinalpropuesta")
 l_puntajemanual = request.querystring("notafinal")
 

 
Response.Write l_evaacuerdo & "<br>" 
Response.Write l_evareunion & "<br>" 
Response.Write l_evaetapa & "<br>" 
Response.Write l_puntaje & "<br>" 
'Response.Write l_puntajemanual & "<br>" 
'Response.Write l_evafecha & "<br>" 
Response.Write l_evaobser & "<br>" 
if trim(l_evaacuerdo)="" or isnull(l_evaacuerdo) then
	l_evaacuerdo=0
end if


'___________________________________________________________________________________

function PasarComaAPunto(valor)
	dim l_numero
	dim l_ubicacion
	dim l_entero
	dim l_decimal
	l_numero = trim(valor)
	l_ubicacion = InStr(l_numero, ",")
	if l_ubicacion > 1 then
		l_ubicacion = l_ubicacion  - 1
		l_entero = left(l_numero, l_ubicacion)
		l_ubicacion = l_ubicacion  + 1
		l_decimal = right(l_numero, (len(l_numero) - l_ubicacion))
    	l_numero = l_entero & "." & l_decimal
    	PasarComaAPunto = l_numero
    else
		PasarComaAPunto = valor
	end if
end function	


'BODY===============================================================================


if l_evaetapa=3 then ' si es cierre de evaluacion, Buscar el resto de los evldrnro
	Set l_rs = Server.CreateObject("ADODB.RecordSet")
	l_sql = "SELECT resto.evldrnro,evadetevldor.evacabnro,evadetevldor.evatevnro  "
	l_sql = l_sql & " FROM  evadetevldor "
	l_sql = l_sql & " INNER JOIN evadetevldor resto ON resto.evacabnro=evadetevldor.evacabnro"
	l_sql = l_sql & "   AND resto.evaseccnro =  evadetevldor.evaseccnro "
	l_sql = l_sql & " WHERE evadetevldor.evldrnro =  " & l_evldrnro
	rsOpen l_rs, cn, l_sql, 0
	l_lista=0
	do while not l_rs.eof
		l_lista= l_lista & "," & l_rs("evldrnro")
		l_evacabnro= l_rs("evacabnro")
		l_evatevnro= l_rs("evatevnro")
		l_rs.Movenext
	loop
	l_rs.close
	set l_rs=nothing
end if
Response.Write l_evatevnro & "<br>" 

'BODY ----------------------------------------------------------

Set l_rs = Server.CreateObject("ADODB.RecordSet")
l_sql = "SELECT *  "
l_sql = l_sql & " FROM  evacierre "
l_sql = l_sql & " WHERE evacierre.evldrnro =  " & l_evldrnro
l_sql = l_sql & "   AND evacierre.evaetapa =  " & l_evaetapa
rsOpen l_rs, cn, l_sql, 0
if l_rs.EOF then
	l_sql = "INSERT INTO evacierre "
	l_sql = l_sql & "(evldrnro, evafecha, evaetapa, evareunion, evaacuerdo,evaobser) "
	l_sql = l_sql & " VALUES (" & l_evldrnro & "," 
	l_sql = l_sql               & cambiafecha(l_evafecha,"YMD",false) & "," 
	l_sql = l_sql               & l_evaetapa   & "," 
	l_sql = l_sql               & l_evareunion & "," 
	l_sql = l_sql               & l_evaacuerdo & ",'" 
	l_sql = l_sql               & l_evaobser   & "')"
	set l_cm = Server.CreateObject("ADODB.Command")  
	l_cm.activeconnection = Cn
	l_cm.CommandText = l_sql
	cmExecute l_cm, l_sql, 0
	
	Response.Write (l_sql) & "<br>"	
else
	l_sql = "UPDATE evacierre SET "
	l_sql = l_sql & " evaobser  = '" & trim(l_evaobser) & "',"
	l_sql = l_sql & " evareunion=  " & l_evareunion & ","
	l_sql = l_sql & " evaacuerdo=  " & l_evaacuerdo & ","
	l_sql = l_sql & " evafecha  =  " & cambiafecha(l_evafecha,"YMD",false) & ""
	l_sql = l_sql & " WHERE evacierre.evldrnro = "  & l_evldrnro
	set l_cm = Server.CreateObject("ADODB.Command")  
	l_cm.activeconnection = Cn
	l_cm.CommandText = l_sql
	cmExecute l_cm, l_sql, 0
	response.write(l_sql)
	if l_evaetapa=3 then
		l_sql = "UPDATE evacierre SET "
		l_sql = l_sql & " evareunion=  " & l_evareunion & ","
		l_sql = l_sql & " evaacuerdo=  " & l_evaacuerdo & ","
		l_sql = l_sql & " evafecha  =  " & cambiafecha(l_evafecha,"YMD",false) & ""
		l_sql = l_sql & " WHERE evacierre.evldrnro IN ("  & l_lista & ")"
		set l_cm = Server.CreateObject("ADODB.Command")  
		l_cm.activeconnection = Cn
		l_cm.CommandText = l_sql
		cmExecute l_cm, l_sql, 0
		response.write(l_sql)
		
		' si es el garante actualizo puntaje manual
		l_sql = "UPDATE evacab SET "
		l_sql = l_sql & " puntaje=  " & PasarComaaPunto(l_puntaje) 
		if l_evatevnro <> cevaluador and l_evatevnro <> cautoevaluador then
		l_sql = l_sql & " , puntajemanual=  " & PasarComaaPunto(l_puntajemanual)
		l_sql = l_sql & " , puntajeevldrnro=  " & l_evldrnro 
		end if
		l_sql = l_sql & " WHERE evacab.evacabnro = "  & l_evacabnro
		set l_cm = Server.CreateObject("ADODB.Command")  
		l_cm.activeconnection = Cn
		l_cm.CommandText = l_sql
		cmExecute l_cm, l_sql, 0
	end if
	Response.Write (l_sql) & "<br>"	
end if
l_rs.Close
set l_rs=nothing

cn.close
set cn=nothing
%>

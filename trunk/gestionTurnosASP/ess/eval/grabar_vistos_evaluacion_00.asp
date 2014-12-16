<% Option Explicit %>
<!--#include virtual="/serviciolocal/shared/inc/sec.inc"-->
<!--#include virtual="/serviciolocal/shared/inc/const.inc"-->
<!--#include virtual="/serviciolocal/shared/db/conn_db.inc"-->
<!--#include virtual="/serviciolocal/shared/inc/fecha.inc"-->
<% 
' variables
' parametros de entrada ----------------------------------------
  Dim l_evldrnro
  Dim l_visdesc
  Dim l_visfecha
  Dim l_visorden
  dim l_puntaje
  dim l_tipo

'locales
 dim	l_evacabnro 
 dim	l_evatevnro 
    
' variables de base de datos ------------------------------------
  Dim l_cm
  Dim l_sql
  Dim l_rs
    
' parametros de entrada
  l_visdesc  = left(trim(request.querystring("visdesc")),200)
  l_evldrnro = request.querystring("evldrnro")
  l_visfecha = request.querystring("visfecha")
  l_visorden = request.querystring("visorden")
  l_puntaje  = request.querystring("puntaje")
  l_tipo     = request.querystring("tipo")

  
  function PasarPuntoAComa(valor)
	dim l_numero
	dim l_ubicacion
	dim l_entero
	dim l_decimal
	l_numero = trim(valor)
	l_ubicacion = InStr(l_numero, ".")
	if l_ubicacion > 1 then
		l_ubicacion = l_ubicacion  - 1
		l_entero = left(l_numero, l_ubicacion)
		l_ubicacion = l_ubicacion  + 1
		l_decimal = right(l_numero, (len(l_numero) - l_ubicacion))
    	l_numero = l_entero & "," & l_decimal
    	PasarPuntoAComa = l_numero
    else
		PasarPuntoAComa = valor
	end if
end function	

'BODY ----------------------------------------------------------

'buscar la evacab
 Set l_rs = Server.CreateObject("ADODB.RecordSet")
 l_sql = "SELECT evacabnro, evatevnro  "
 l_sql = l_sql & " FROM  evadetevldor "
 l_sql = l_sql & " WHERE evldrnro   = " & l_evldrnro
 rsOpen l_rs, cn, l_sql, 0
 if not l_rs.EOF then
	l_evacabnro = l_rs("evacabnro")
	l_evatevnro = l_rs("evatevnro")
 end if
 l_rs.close
 set l_rs=nothing
 
if l_tipo="M" then
	l_sql = "UPDATE evavistos SET "
	l_sql = l_sql & " visdesc = '" & trim(l_visdesc) & "',"
	l_sql = l_sql & " visfecha = " & cambiafecha(l_visfecha,"YMD",false) & ""
	l_sql = l_sql & " WHERE evavistos.evldrnro = "  & l_evldrnro
	set l_cm = Server.CreateObject("ADODB.Command")  
	l_cm.activeconnection = Cn
	l_cm.CommandText = l_sql
	cmExecute l_cm, l_sql, 0
	Response.Write (l_sql) & "<br>"	
	
	'grabar puntaje
	 if trim(l_puntaje)<>"" then
	' esto para cuando se usa evapuntaje !!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!
		'Set l_rs = Server.CreateObject("ADODB.RecordSet")
		'l_sql = "SELECT puntaje  "
		'l_sql = l_sql & " FROM  evapuntaje "
		'l_sql = l_sql & " WHERE evacabnro   = " & l_evacabnro
		'l_sql = l_sql & " AND	 evatevnro   = " & l_evatevnro
		'rsOpen l_rs, cn, l_sql, 0
		'Response.Write (l_sql) & "<br>"	
		'if not l_rs.EOF then
		'	l_rs.close
		'	set l_rs=nothing
		'	l_sql = "UPDATE evapuntaje SET "
		'	l_sql = l_sql & " puntajemanual = " & l_puntaje
		'	l_sql = l_sql & " WHERE evacabnro = "  & l_evacabnro
		'	l_sql = l_sql & " AND   evatevnro = "  & l_evatevnro
		'	set l_cm = Server.CreateObject("ADODB.Command")  
		'	l_cm.activeconnection = Cn
		'	l_cm.CommandText = l_sql
		'	cmExecute l_cm, l_sql, 0
		'	Response.Write (l_sql) & "<br>"	
		'else	
		'	l_sql = "INSERT INTO evapuntaje "
		'	l_sql = l_sql & "(evacabnro,evatevnro,puntajemanual) "
		'	l_sql = l_sql & " VALUES (" & l_evacabnro & "," 
		'	l_sql = l_sql               & evatevnro & "," & l_puntaje & ")"
		'	set l_cm = Server.CreateObject("ADODB.Command")  
		'	l_cm.activeconnection = Cn
		'	l_cm.CommandText = l_sql
		'	cmExecute l_cm, l_sql, 0
		'	Response.Write (l_sql) & "<br>"	
		'end if
		l_sql = "UPDATE evacab SET "
		l_sql = l_sql & " puntajemanual = " & l_puntaje &","
		l_sql = l_sql & " puntajeevldrnro = " & l_evldrnro
		l_sql = l_sql & " WHERE evacabnro = "  & l_evacabnro
		set l_cm = Server.CreateObject("ADODB.Command")  
		l_cm.activeconnection = Cn
		l_cm.CommandText = l_sql
		cmExecute l_cm, l_sql, 0
		Response.Write (l_sql) & "<br>"	
	end if	
else ' es ALTA 
	l_sql = "INSERT INTO evavistos "
	l_sql = l_sql & "(evldrnro,visfecha,visorden, visdesc) "
	l_sql = l_sql & " VALUES (" & l_evldrnro & "," 
	l_sql = l_sql               & cambiafecha(l_visfecha,"YMD",false) & ",1,'" & l_visdesc & "')"
	set l_cm = Server.CreateObject("ADODB.Command")  
	l_cm.activeconnection = Cn
	l_cm.CommandText = l_sql
	cmExecute l_cm, l_sql, 0

	if trim(l_puntaje)<>"" then
' esto para cuando se usa evapuntaje !!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!
'		l_sql = "INSERT INTO evapuntaje "
'		l_sql = l_sql & "(evacabnro,evatevnro,puntajemanual) "
'		l_sql = l_sql & " VALUES (" & l_evacabnro & "," 
'		l_sql = l_sql               & evatevnro & "," & l_puntaje & ")"
'		set l_cm = Server.CreateObject("ADODB.Command")  
'		l_cm.activeconnection = Cn
'		l_cm.CommandText = l_sql
'		cmExecute l_cm, l_sql, 0
		l_sql = "UPDATE evacab SET "
		l_sql = l_sql & " puntajemanual = " & l_puntaje &","
		l_sql = l_sql & " puntajeevldrnro = " & l_evldrnro
		l_sql = l_sql & " WHERE evacabnro = "  & l_evacabnro
		set l_cm = Server.CreateObject("ADODB.Command")  
		l_cm.activeconnection = Cn
		l_cm.CommandText = l_sql
		cmExecute l_cm, l_sql, 0
	end if

end if

%>

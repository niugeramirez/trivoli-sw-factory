<%

'----------------------------------------------------------------------------------------------------------
'FUNCION: incluye un archivo de estilos en la pagina web 
sub incluirEstilo(filesystem,archivo,nombreestilo)
  Dim pathestilo
  Dim archestilo
  Dim str

  pathestilo = Server.MapPath("/serviciolocal/shared/css/") 
  pathestilo = replace(pathestilo,"\","\\")  

  Set archestilo = filesystem.OpenTextFile(pathestilo & "\\" & nombreestilo,1,false) 

  guardarTexto archivo,"<style>"
  Do While archestilo.AtEndOfStream <> true
     str=archestilo.readline
	 guardarTexto archivo,str
  Loop 
  guardarTexto archivo,"</style>"
  
  archestilo.close

end sub 'incluirEstilo(archivo)

'----------------------------------------------------------------------------------------------------------
'FUNCION: genera un numero unico
Function genguid     
     Dim Guid
    guid = server.createobject("scriptlet.typelib").guid
    guid=Left(guid,instr(guid,"}"))
    guid=replace(guid,"{","")
    guid=replace(guid,"}","")
    guid=replace(guid,"-","")
    genguid= guid
    set guid=nothing	
end function

'----------------------------------------------------------------------------------------------------------
'FUNCION: guarda un texto en un archivo 
sub guardarTexto(archivo,texto)
  archivo.write texto
end sub 

'----------------------------------------------------------------------------------------------------------
'FUNCION: guarda una linea en un archivo 
sub guardarLinea(archivo,texto)
  archivo.writeline texto
end sub 

'----------------------------------------------------------------------------------------------------------
'FUNCION: crear el archivo y el proceso para enviar un mail
sub enviarMail(filesystem,archivo,asunto,msgbody,destino)

Dim m_sql
Dim m_rs

Set m_rs = Server.CreateObject("ADODB.RecordSet")

m_sql = " SELECT * FROM conf_email WHERE cfgemailest = -1 "

rsOpen m_rs, cn, m_sql, 0 

if not m_rs.eof then

    Dim m_host
	Dim m_from
	Dim m_port
	
    m_host = m_rs("cfgemailhost")
	m_from = m_rs("cfgemailfrom")
	m_port = m_rs("cfgemailport")
	
	m_rs.close

	Dim archlog
	
	Dim archpath
	Dim nombrearch
	
	nombrearch = "mail_" & genGuid & ".msg"
	
	archpath = Server.MapPath("/serviciolocal/log/")
	archpath = replace(archpath,"\","\\")
	archpath = archpath & "\\" & nombrearch
	
	Set arch = fs.CreateTextFile(archpath, True)
	
	guardarLinea arch,"[MailMessage]"
	guardarLinea arch,"FromName=RHPro Msg"
	guardarLinea arch,"FromAddress=" & m_from
	guardarLinea arch,"Subject=" & asunto
	guardarLinea arch,"Body1=" & msgbody
	guardarLinea arch,"Attachment=" & archivo
	guardarLinea arch,"Recipients=" & destino
	guardarLinea arch,"Host=" & m_host
	guardarLinea arch,"Port=" & m_port
	
	arch.close
	
	Dim m_cm
	set m_cm = Server.CreateObject("ADODB.Command")
	m_cm.activeconnection = cn
	
	
	'Ingreso el proceso a la tabla
	m_sql = generarSQLProc(25,Date,Date,"")
	cmExecute m_cm, m_sql, 0
	
	
else
   	m_rs.close	
end if

end sub 

' ------------------------------------------------------------------------------------------------------------------
' generarSQLProc(tipoPorc,desde,hasta) :
' parametros:
'    * tipoProc     : tipo de proceso a insertar
'    * desde, hasta : fechas desde y hasta sin formatear
'    * parametros   : parametros opcionales del procedimiento
' ------------------------------------------------------------------------------------------------------------------
function generarSQLProc(tipoPorc,desde,hasta,parametros)

Dim l_id
Dim l_hora
Dim l_dia
Dim l_sqlp
Dim l_desde_sql
Dim l_hasta_sql

l_desde_sql = cambiafecha(desde,"YMD",true)
l_hasta_sql = cambiafecha(hasta,"YMD",true)

l_id   = Session("Username")
l_hora = mid(time,1,8)
l_dia  = cambiafecha(Date,"YMD",true)

l_sqlp =          " INSERT INTO batch_proceso "
l_sqlp = l_sqlp & " (btprcnro, bprcfecha, iduser, bprchora, bprcfecdesde, bprcfechasta, bprcparam, "
l_sqlp = l_sqlp & " bprcestado, bprcprogreso, bprcfecfin, bprchorafin, bprctiempo, empnro, bprcempleados,bprcurgente) "
l_sqlp = l_sqlp & " VALUES (" & tipoPorc & "," & l_dia & ", '"& l_id &"','"& l_hora &"' "
l_sqlp = l_sqlp & " , " & l_desde_sql & ", " & l_hasta_sql
l_sqlp = l_sqlp & " , '" & parametros & "', 'Pendiente', null , null, null, null, 0, null,0)"


generarSQLProc = l_sqlp

end function 'generarSQLProc()

%>
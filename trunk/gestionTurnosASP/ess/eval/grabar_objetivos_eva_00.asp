<% Option Explicit %>
<!--#include virtual="/serviciolocal/shared/inc/sec.inc"-->
<!--#include virtual="/serviciolocal/shared/inc/const.inc"-->
<!--#include virtual="/serviciolocal/shared/db/conn_db.inc"-->
<!--#include virtual="/serviciolocal/shared/inc/fecha.inc"-->
<!--#include virtual="/serviciolocal/shared/inc/sqls.inc"-->
<!--#include virtual="/serviciolocal/shared/inc/const_eva.inc"-->
<% 
'=====================================================================================
'Archivo  	: grabar_objetivos_eva_00.asp
'Objetivo 	: grabar de objetivos de evaluacion
'Fecha	  	: 06-05-2004
'Autor	  	: CCRossi
'Modificacion: 26-04-2005 - LA - Si es la priemra vez que se define Objs, refrescar el estado de seccion(p/deloitte)
'			   14-06-2006 - LA -  agregar deletes de tablas relacionada a objetivo (evaluaobj).
'			   12-12-2006 - LA - ir modularizando (sacar todas las def de ll_rs, l_rs1 y cn)
'								  Agregar ABM para el nuevo  tipo de seccion PL (Obj con Plan de desarrollo)
'			   15-06-2006 - LA - incorporar todas las mejoras del modulo
'=====================================================================================

on error goto 0

' variables
' parametros de entrada ----------------------------------------
  Dim l_evldrnro
  Dim l_evaobjnro
  Dim l_evaobjdext
  Dim l_evaobjformed
  Dim l_evaperfijo
  Dim l_evaobjplan
Dim l_evaobjresu
Dim l_evaobjfecha 
  
  
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
  
  Dim primerObjetivo
  
' parametros de entrada
l_tipo		 = request.querystring("tipo")
l_evaobjnro	 = request.querystring("evaobjnro")
l_evldrnro	 = request.querystring("evldrnro")
l_evaobjdext  = request.querystring("evaobjdext")
l_evaobjformed= request.querystring("evaobjformed")
l_evaperfijo  = request.querystring("evapernro")
l_evaobjplan  = request.Form("evaobjplan")
l_evaobjresu  = request.Form("evaobjresu")
l_evaobjfecha = request.Form("evaobjfecha")
 
Set l_rs = Server.CreateObject("ADODB.RecordSet")
Set l_rs1 = Server.CreateObject("ADODB.RecordSet")
set l_cm = Server.CreateObject("ADODB.Command")  

'response.write l_evaobjplan

' _________________________________________________
if l_evaobjfecha <> "" then
l_evaobjfecha 	= cambiafecha(l_evaobjfecha,"",false)  
else
l_evaobjfecha="NULL"
end if 
l_evaobjdext    = left(trim(l_evaobjdext),250)
l_evaobjformed  = left(trim(l_evaobjformed),250)
l_evaobjplan 	= left(trim(l_evaobjplan),1500)
l_evaobjresu	= left(trim(l_evaobjresu),1500)



primerObjetivo="NO"
l_sql = "SELECT evaluaobj.evaobjnro FROM evaobjetivo "
l_sql = l_sql & " INNER JOIN evaluaobj ON evaluaobj.evaobjnro = evaobjetivo.evaobjnro"
l_sql = l_sql & " WHERE evaluaobj.evldrnro =" & l_evldrnro
rsOpen l_rs, cn, l_sql, 0
if l_rs.eof then
	primerObjetivo = "SI"
end if
l_rs.Close



'BODY ----------------------------------------------------------

' borrar registros de la tabla relacionada a evaobjetivo
if l_tipo="B" or l_tipo="BPL" then
	l_sql = "SELECT * FROM evaluaobj WHERE evaobjnro = " & l_evaobjnro
		'l_sql = l_sql & " AND   evldrnro  = " & l_evldrnro, este no porque crea registros de evaluobj para otros evaluadores y otras secciones de tipoobj
	rsOpen l_rs1, cn, l_sql, 0
	do while not l_rs1.eof 
		l_sql = "DELETE FROM evaluaobj "
		l_sql = l_sql & " WHERE evaluaobj.evaobjnro= " & l_evaobjnro & " AND evaluaobj.evldrnro =" & l_rs1("evldrnro")
		l_cm.activeconnection = Cn
		l_cm.CommandText = l_sql
		cmExecute l_cm, l_sql, 0
		
	l_rs1.MoveNext
	loop
	
	l_rs1.Close
end if

' borra regsitros con informacion adicional al objetivo a borrar
if l_tipo="BPL" then
	l_sql = "DELETE FROM evaobjplan "
	l_sql = l_sql & " WHERE evaobjplan.evaobjnro=" & l_evaobjnro
	l_cm.activeconnection = Cn
	l_cm.CommandText = l_sql
	cmExecute l_cm, l_sql, 0
end if

' 
select case l_tipo
Case "A", "PL":
		l_sql= "insert into evaobjetivo (evaobjdext,evaobjformed,evaperfijo) "
		l_sql = l_sql & " values ('" & trim(l_evaobjdext) & "','" & trim(l_evaobjformed) & "'," & l_evaperfijo &")"
case "M","MPL":
		l_sql = "UPDATE evaobjetivo SET "
		l_sql = l_sql & " evaobjdext   = '" & trim(l_evaobjdext) & "',"
		l_sql = l_sql & " evaobjformed = '" & trim(l_evaobjformed) & "',"
		l_sql = l_sql & " evaperfijo   = "  & l_evaperfijo & " "
		l_sql = l_sql & " WHERE evaobjetivo.evaobjnro = "  & l_evaobjnro
case "B","BPL":
		l_sql = "DELETE from evaobjetivo where evaobjetivo.evaobjnro = "  & l_evaobjnro
case "E":
		dim l_evatrnro
		dim l_evapernroeva
		dim l_evapereva
		l_evatrnro = Request.QueryString("evatrnro")
		l_evapernroeva = Request.QueryString("evapernro")
		l_evapereva = request.querystring("evldrnro")
		
		l_sql = "UPDATE evaobjetivo SET "
		l_sql = l_sql & " evaobjdext    = '" & trim(l_evaobjdext) & "',"
		l_sql = l_sql & " evaobjformed  = '" & trim(l_evaobjformed) & "',"
		l_sql = l_sql & " evapernroeva  = "  & l_evapernroeva & ","
		l_sql = l_sql & " evapereva     = "  & l_evapereva & " "
		'l_sql = l_sql & " evatrnro		= "  & l_evatrnro & " "
		l_sql = l_sql & " WHERE evaobjetivo.evaobjnro = "  & l_evaobjnro
end select

l_cm.activeconnection = Cn
l_cm.CommandText = l_sql
cmExecute l_cm, l_sql, 0


if trim(l_evaobjnro)="" or l_evaobjnro=0 then
	l_sql = fsql_seqvalue("evaobjnro","evaobjetivo")
	rsOpen l_rs, cn, l_sql, 0
	if not l_rs.eof then
		l_evaobjnro=l_rs("evaobjnro")
	end if	
	l_rs.Close
end if


' Guardar informacion adicional ____________________________________
select case l_tipo  ' Pl 
Case "PL":
	l_sql = "SELECT * "
	l_sql = l_sql & " FROM evaobjplan "
	l_sql = l_sql & " WHERE evaobjplan.evaobjnro="& l_evaobjnro
	rsOpen l_rs, cn, l_sql, 0
	if l_rs.eof then
		l_sql= "INSERT INTO evaobjplan(evaobjnro,evaobjplan,evaobjresu,evaobjfecha) "
		l_sql = l_sql & " VALUES("& l_evaobjnro &", '"& l_evaobjplan & "','"& l_evaobjresu & "',"& l_evaobjfecha  &")"
	end if
	
	l_cm.activeconnection = Cn
	l_cm.CommandText = l_sql
	cmExecute l_cm, l_sql, 0
	
	l_rs.Close
	
case "MPL":
	l_sql= "UPDATE evaobjplan SET "
	l_sql = l_sql & " evaobjplan='" & l_evaobjplan & "',"
	l_sql = l_sql & " evaobjresu='" & l_evaobjresu & "',"
	l_sql = l_sql & " evaobjfecha=" & l_evaobjfecha  
	l_sql = l_sql & " WHERE evaobjnro ="& l_evaobjnro
	
	l_cm.activeconnection = Cn
	l_cm.CommandText = l_sql
	cmExecute l_cm, l_sql, 0
	
end select


' CREAR REGISTROS DE LA RELACION  ____________________________________
select case l_tipo
Case "A", "PL":
	l_sql = "SELECT * FROM evaluaobj WHERE evaobjnro = " & l_evaobjnro
	l_sql = l_sql & " AND   evldrnro  = " & l_evldrnro
	rsOpen l_rs1, cn, l_sql, 0
	if  l_rs1.eof then
		l_sql= "insert into evaluaobj (evldrnro,evaobjnro) "
		l_sql = l_sql & " values (" & l_evldrnro & "," & l_evaobjnro &")"
		l_cm.activeconnection = Cn
		l_cm.CommandText = l_sql
		cmExecute l_cm, l_sql, 0
	end if	
	l_rs1.Close
	
Case "E":
	l_sql = "SELECT * FROM evaluaobj WHERE evaobjnro = " & l_evaobjnro
	l_sql = l_sql & " AND   evldrnro  = " & l_evldrnro
	rsOpen l_rs1, cn, l_sql, 0
	if  l_rs1.eof then
		l_sql= "insert into evaluaobj (evldrnro,evaobjnro,evatrnro) "
		l_sql = l_sql & " values (" & l_evldrnro & "," & l_evaobjnro &","& l_evatrnro&")"
		l_cm.activeconnection = Cn
		l_cm.CommandText = l_sql
		cmExecute l_cm, l_sql, 0
	else	
		l_sql= "UPDATE evaluaobj SET evatrnro = " & l_evatrnro
		l_sql = l_sql & " WHERE evldrnro="& l_evldrnro 
		l_sql = l_sql & " AND evaobjnro ="& l_evaobjnro
		l_cm.activeconnection = Cn
		l_cm.CommandText = l_sql
		cmExecute l_cm, l_sql, 0
	end if	
	l_rs1.Close
end select



if l_tipo <> "B" and l_tipo <> "BPL" then
	
	l_sql = "SELECT * FROM evadetevldor WHERE evldrnro = "& l_evldrnro
	rsOpen l_rs, cn, l_sql, 0
	if not l_rs.eof then
		l_evacabnro =l_rs("evacabnro")
		l_evatevnro =l_rs("evatevnro")
		l_evaluador =l_rs("evaluador")
	end if	
	l_rs.Close

	l_sql = "SELECT evldrnro FROM evadetevldor INNER JOIN evasecc ON evadetevldor.evaseccnro = evasecc.evaseccnro "
	l_sql = l_sql & " INNER JOIN evatiposecc ON evasecc.tipsecnro = evatiposecc.tipsecnro "
	l_sql = l_sql & " WHERE evacabnro = " & l_evacabnro
	l_sql = l_sql & " AND   evatevnro = " & l_evatevnro
	l_sql = l_sql & " AND   evaluador = " & l_evaluador
	l_sql = l_sql & " AND   evldrnro  <> " & l_evldrnro
	l_sql = l_sql & " AND   tipsecobj=-1" 
	rsOpen l_rs, cn, l_sql, 0
	
	do while not l_rs.eof 
		l_sql = "SELECT * FROM evaluaobj WHERE evaobjnro = " & l_evaobjnro
		l_sql = l_sql & " AND   evldrnro  = " & l_rs("evldrnro")
		rsOpen l_rs1, cn, l_sql, 0
		if  l_rs1.eof then
			l_sql= "insert into evaluaobj (evldrnro,evaobjnro) "
			l_sql = l_sql & " values (" & l_rs("evldrnro") & "," & l_evaobjnro &")"
			l_cm.activeconnection = Cn
			l_cm.CommandText = l_sql
			cmExecute l_cm, l_sql, 0
		end if
		
		l_rs1.Close
		l_rs.MoveNext
	loop
	l_rs.Close

	if cmismo_objetivo=1 then ' todos los evaluadores evaluan los mismo objetivos para el mismo EVALUADO
	'crear los evaluaobj para LOS OTROS Evaluadores 
	'busco los evldrnro de los otros evatevnro
	l_sql = "SELECT evldrnro FROM evadetevldor INNER JOIN evasecc ON evadetevldor.evaseccnro = evasecc.evaseccnro "
	l_sql = l_sql & " INNER JOIN evatiposecc ON evasecc.tipsecnro = evatiposecc.tipsecnro "
	l_sql = l_sql & " WHERE evacabnro = " & l_evacabnro
	l_sql = l_sql & " AND   evatevnro  <> " & l_evatevnro
	l_sql = l_sql & " AND   tipsecobj=-1" 
	rsOpen l_rs, cn, l_sql, 0
	
	do while not l_rs.eof 
		l_evldrnro= l_rs("evldrnro")
		'busco el objetivo si ya existe para este evldrnro
		l_sql = "SELECT * FROM evaluaobj WHERE evaobjnro = " & l_evaobjnro
		l_sql = l_sql & " AND   evldrnro  = " & l_evldrnro
		rsOpen l_rs1, cn, l_sql, 0
		if  l_rs1.eof then
			l_sql= "insert into evaluaobj (evldrnro,evaobjnro) "
			l_sql = l_sql & " values (" & l_evldrnro & "," & l_evaobjnro &")"
			l_cm.activeconnection = Cn
			l_cm.CommandText = l_sql
			cmExecute l_cm, l_sql, 0
		end if
		
		l_rs1.Close
		l_rs.MoveNext
	loop
	l_rs.Close
	end if
end if ' <> "B"

cn.close
set cn=nothing

 '  ver para que habilite la seccion T , en el primer def obj  ---'cint(cdeloitte)= -1 and
if primerObjetivo="SI" then
	response.write "<script> parent.parent.estado.location.reload();</script>"
end if
response.write "<script> parent.location.reload(); </script>"
%>

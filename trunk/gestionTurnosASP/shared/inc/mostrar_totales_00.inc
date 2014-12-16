<%
' rhpro x2--------------------------------

dim l_Error
Dim l_tipseccalif
Dim l_evacabnro
Dim l_rs_oblig
Set l_rs = Server.CreateObject("ADODB.RecordSet")

l_Error = 0

if trim(l_evacabnro)="" or isnull(l_evacabnro) then
' Buscar EVACAB y el tipo de evaluador actual
l_sql = "SELECT evacab.evacabnro "
l_sql = l_sql & " FROM evadetevldor "
l_sql = l_sql & " INNER JOIN evacab  ON evadetevldor.evacabnro=evacab.evacabnro "
l_sql = l_sql & " WHERE evadetevldor.evldrnro=" & l_evldrnro 
rsOpen l_rs, cn, l_sql, 0
if not l_rs.eof then
	l_evacabnro = l_rs("evacabnro")    'NUEVO
end if	
l_rs.Close
end if

'Busca el codigo del tipo de seccion competencias
l_sql = "SELECT evasecc.tipsecnro "
l_sql = l_sql & " FROM evasecc "
l_sql = l_sql & " INNER JOIN evaseccfactor ON evasecc.evaseccnro = evaseccfactor.evaseccnro "
l_sql = l_sql & " INNER JOIN evadetevldor  ON evasecc.evaseccnro  = evadetevldor.evaseccnro "
l_sql = l_sql & " WHERE evadetevldor.evacabnro=" & l_evacabnro
rsOpen l_rs, cn, l_sql, 0 
if not l_rs.eof then
	l_tipseccalif = l_rs("tipsecnro")
else
	l_tipseccalif = "0"
end if
l_rs.Close

' Buscar el evaseccnro para tiposecnro= l_tipseccalif (competencias)
l_sql = "SELECT evadet.evaseccnro "
l_sql = l_sql & " FROM evadetevldor "
l_sql = l_sql & " INNER JOIN evacab  ON evadetevldor.evacabnro=evacab.evacabnro "
l_sql = l_sql & " INNER JOIN evadet  ON evadet.evacabnro      =evacab.evacabnro "
l_sql = l_sql & " INNER JOIN evasecc ON evasecc.evaseccnro    =evadet.evaseccnro "
l_sql = l_sql & "		AND evasecc.tipsecnro = "& l_tipseccalif
l_sql = l_sql & " WHERE evadetevldor.evldrnro=" & l_evldrnro 
rsOpen l_rs, cn, l_sql, 0
if not l_rs.eof then
	l_evaseccnro= l_rs("evaseccnro")
end if	
l_rs.Close

'Buscar tipo evaluador y el evldrnro de la seccion de competencias, para el eveluador actual
Set l_rs = Server.CreateObject("ADODB.RecordSet")
l_sql = "SELECT evaluador.evldrnro, evaluador.evatevnro "
l_sql = l_sql & " FROM evadetevldor "
l_sql = l_sql & " INNER JOIN evacab				    ON evadetevldor.evacabnro=evacab.evacabnro "
l_sql = l_sql & " INNER JOIN evadetevldor evaluador ON evacab.evacabnro = evaluador.evacabnro "
l_sql = l_sql & "	     AND evaluador.evatevnro = evadetevldor.evatevnro "
l_sql = l_sql & " WHERE evadetevldor.evldrnro  = " & l_evldrnro  
l_sql = l_sql & "   AND evacab.empleado		   = " & l_empleado
l_sql = l_sql & "   AND evaluador.evaseccnro   = " & l_evaseccnro 'seccion de comptencias
rsOpen l_rs, cn, l_sql, 0
if not l_rs.eof then
	l_evatevnro = l_rs("evatevnro")
	l_evldrnro  = l_rs("evldrnro")
end if	
l_rs.Close

'=========================================================================================
' buscar potencial
Set l_rs = Server.CreateObject("ADODB.RecordSet")
l_sql = "SELECT evaluadores.evldrnro "
l_sql = l_sql & " FROM evadetevldor "
l_sql = l_sql & " INNER JOIN evacab		  ON evadetevldor.evacabnro=evacab.evacabnro "
l_sql = l_sql & " INNER JOIN evadetevldor evaluadores ON evacab.evacabnro = evaluadores.evacabnro "
l_sql = l_sql & " INNER JOIN evatipevalua ON evatipevalua.evatevnro = evaluadores.evatevnro "
l_sql = l_sql & "        AND evatipevalua.evatevpot = -1 " 
l_sql = l_sql & " WHERE evadetevldor.evldrnro=" & l_evldrnro 
l_sql = l_sql & "   and evacab.empleado="& l_empleado
l_sql = l_sql & "   and evaluadores.evaseccnro=" & l_evaseccnro 
rsOpen l_rs, cn, l_sql, 0
if not l_rs.eof then
	l_potencial = l_rs("evldrnro")
else
	l_rs.close
	set l_rs=nothing
	Set l_rs = Server.CreateObject("ADODB.RecordSet")
	l_sql = " SELECT evaresultado.evafacnro "
	l_sql = l_sql & " FROM  evaresultado "
	l_sql = l_sql & " INNER JOIN evatipresu    ON evatipresu.evatrnro = evaresultado.evatrnro   "
	l_sql = l_sql & " INNER JOIN evaseccfactor ON evaseccfactor.evafacnro = evaresultado.evafacnro "
	l_sql = l_sql & " INNER JOIN evafactor	 ON evaseccfactor.evafacnro = evafactor.evafacnro "
	l_sql = l_sql & "		 AND   evafactor.evafacpot = -1 " 
	l_sql = l_sql & " INNER JOIN evadetevldor	 ON evadetevldor.evldrnro = evaresultado.evldrnro "
	l_sql = l_sql & "		 AND   evadetevldor.evacabnro =  " &l_evacabnro
	l_sql = l_sql & " WHERE evaseccfactor.evaseccnro =  " & l_evaseccnro 
	rsOpen l_rs, cn, l_sql, 0
	if l_rs.eof then 
		Response.Write("<script>alert('No se ha definido un Rol Potencial o Competencia Potencial.')</script>")
		l_Error = -1
	end if
	l_rs.close
	set l_rs=nothing
end if	


'=========================================================================================
' Calcular Suma Total
 Set l_rs = Server.CreateObject("ADODB.RecordSet")
 l_sql = " SELECT SUM(evatrvalor) AS suma "
 l_sql = l_sql & " FROM  evaresultado "
 l_sql = l_sql & " INNER JOIN evatipresu    ON evatipresu.evatrnro = evaresultado.evatrnro   "
 l_sql = l_sql & " INNER JOIN evaseccfactor ON evaseccfactor.evafacnro = evaresultado.evafacnro "
 l_sql = l_sql & " INNER JOIN evafactor		ON evafactor.evafacnro = evaresultado.evafacnro "
 l_sql = l_sql & "		AND evafactor.evafacpot = 0 "
 l_sql = l_sql & " WHERE evaseccfactor.evaseccnro = " & l_evaseccnro 
 l_sql = l_sql & " AND   evaresultado.evldrnro    = " & l_evldrnro
 rsOpen l_rs, cn, l_sql, 0
 if not l_rs.eof then 
	l_suma = l_rs("suma")
 end if
 l_rs.Close
 'Buscar el maximo valor    
  l_sql = "SELECT MAX(evatrvalor) AS maximo "
  l_sql = l_sql & " FROM  evatipresu "
  l_sql = l_sql & " INNER JOIN evaresu       ON evaresu.evatrnro = evatipresu.evatrnro "
  l_sql = l_sql & " WHERE evaresu.evaseccnro = " & l_evaseccnro
  Set l_rs = Server.CreateObject("ADODB.RecordSet")
  rsOpen l_rs, cn, l_sql, 0
  if not l_rs.eof then 
	l_maximo = l_rs("maximo")
  end if
  l_rs.Close

'response.write l_maximo & "<br>"

  l_sql = "SELECT COUNT(evaseccfactor.evafacnro) AS cantidad "
  l_sql = l_sql & " FROM evaseccfactor "
  l_sql = l_sql & " INNER JOIN evafactor ON evafactor.evafacnro = evaseccfactor.evafacnro "
  l_sql = l_sql & "		AND evafactor.evafacpot = 0 "
  l_sql = l_sql & " WHERE evaseccfactor.evaseccnro = " & l_evaseccnro
  l_sql = l_sql & " AND EXISTS "
  l_sql = l_sql & " (SELECT * FROM evaresultado "
  l_sql = l_sql & " INNER JOIN evatipresu ON evatipresu.evatrnro = evaresultado.evatrnro "
  l_sql = l_sql & " AND evatipresu.evatrnro IS NOT NULL "
  l_sql = l_sql & " WHERE evaresultado.evafacnro = evaseccfactor.evafacnro "
  l_sql = l_sql & " AND   evaresultado.evldrnro = " & l_evldrnro
  l_sql = l_sql & ")"
  Set l_rs = Server.CreateObject("ADODB.RecordSet")
  rsOpen l_rs, cn, l_sql, 0
  if not l_rs.eof then 
	l_cantidad = l_rs("cantidad")
  end if
  l_rs.Close
  
'response.write l_cantidad & "<br>"


'fin del include
%>
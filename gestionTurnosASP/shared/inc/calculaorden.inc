<%  
Sub OrdenTurno(l_turnro)
Dim l_orden
Dim l_ordensubturno	
Dim l_subturnro
Dim rsOrden
Dim cmOrden
set rsOrden = Server.CreateObject("ADODB.RecordSet")
set cmOrden = Server.CreateObject("ADODB.Command")	
l_orden = 1
l_ordenSubturno = 0
l_subturnro = 0
l_Sql = "SELECT * FROM gti_subturno, gti_dias " 
l_sql = l_sql & " WHERE turnro = " & l_turnro
l_sql = l_sql & " AND gti_dias.subturnro = gti_subturno.subturnro"
l_sql = l_sql & " ORDER BY gti_dias.diaorden "
rsOpen rsOrden, cn, l_sql, 0 
do until rsOrden.eof
	l_sql = "UPDATE gti_dias SET diaorden =  " & l_orden & " WHERE dianro = " & rsOrden("dianro")
'	response.write l_sql
	cmOrden.activeconnection = Cn
	cmOrden.CommandText = l_sql
	cmExecute cmOrden, l_sql, 0
	l_orden = l_orden + ( rsOrden("diaiteracion") * rsOrden("subturiteracion") )
	rsOrden.MoveNext
loop
rsOrden.close
l_Sql = "SELECT * FROM gti_subturno " 
l_sql = l_sql & " WHERE turnro = " & l_turnro
l_sql = l_sql & " ORDER BY subturorden "
rsOpen rsOrden, cn, l_sql, 0 
l_ordenSubturno = 0
do until rsOrden.eof
	l_ordenSubturno = l_ordenSubturno + 1
	l_sql = "UPDATE gti_subturno SET subturorden =  " & l_ordenSubturno  & " WHERE subturnro = " & rsOrden("subturnro")
	cmOrden.activeconnection = Cn
	cmOrden.CommandText = l_sql
	cmExecute cmOrden, l_sql, 0
	rsOrden.MoveNext
loop
rsOrden.close
end sub

Sub OrdenTurno2(l_turnro)
Dim l_orden
Dim l_ordensubturno	
Dim l_subturnro
Dim rsOrden
Dim cmOrden
set rsOrden = Server.CreateObject("ADODB.RecordSet")
set cmOrden = Server.CreateObject("ADODB.Command")	
l_orden = 1
l_ordenSubturno = 0
l_subturnro = 0
l_Sql = "SELECT * FROM gti_subturno, gti_dias " 
l_sql = l_sql & " WHERE turnro = " & l_turnro
l_sql = l_sql & " AND gti_dias.subturnro = gti_subturno.subturnro"
l_sql = l_sql & " ORDER BY subturorden, gti_dias.diaorden "
rsOpen rsOrden, cn, l_sql, 0 
do until rsOrden.eof
	l_sql = "UPDATE gti_dias SET diaorden =  " & l_orden & " WHERE dianro = " & rsOrden("dianro")
'	response.write l_sql
	cmOrden.activeconnection = Cn
	cmOrden.CommandText = l_sql
	cmExecute cmOrden, l_sql, 0
	l_orden = l_orden + ( rsOrden("diaiteracion") * rsOrden("subturiteracion") )
	rsOrden.MoveNext
loop
rsOrden.close
l_Sql = "SELECT * FROM gti_subturno " 
l_sql = l_sql & " WHERE turnro = " & l_turnro
l_sql = l_sql & " ORDER BY subturorden "
rsOpen rsOrden, cn, l_sql, 0 
l_ordenSubturno = 0
do until rsOrden.eof
	l_ordenSubturno = l_ordenSubturno + 1
	l_sql = "UPDATE gti_subturno SET subturorden =  " & l_ordenSubturno  & " WHERE subturnro = " & rsOrden("subturnro")
	cmOrden.activeconnection = Cn
	cmOrden.CommandText = l_sql
	cmExecute cmOrden, l_sql, 0
	rsOrden.MoveNext
loop
rsOrden.close
end sub


%>	
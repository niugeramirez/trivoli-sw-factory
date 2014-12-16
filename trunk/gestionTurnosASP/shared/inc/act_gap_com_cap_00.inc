<% 

Dim m_rs
Dim m_rs2
'Dim l_sql
'Dim l_ternro
'Dim l_filtro
'Dim l_orden
'Dim l_opc
'Dim l_entnro
'Dim l_hay_datos

Dim terminada
Dim m_cm

'l_ternro = request("ternro")

Set m_rs = Server.CreateObject("ADODB.RecordSet")
Set m_rs2 = Server.CreateObject("ADODB.RecordSet")
set m_cm = Server.CreateObject("ADODB.Command")

' Detecto los gap por competencias del tercero
l_sql = "SELECT evafacnro, falnro, modnro, falorigen, evafacdesabr, falidnro, falfecha, falpendiente, falporcen"
l_sql = l_sql & " FROM cap_falencia "
l_sql = l_sql & " INNER JOIN evafactor ON evafactor.evafacnro = cap_falencia.modnro "
l_sql = l_sql & " WHERE ternro = " & l_ternro 
l_sql = l_sql & " AND falorigen = 7 "
l_sql = l_sql & " ORDER BY falorigen, falnro "

rsOpen m_rs, cn, l_sql, 0
do until m_rs.eof
	terminada = false

	'MANUAL
	l_sql = "SELECT cap_capacita.porcen, cap_capacita.fecha "
	l_sql = l_sql & " FROM cap_capacita "
	l_sql = l_sql & " WHERE (cap_capacita.origen1 = 5) AND " 'MANUAL
	l_sql = l_sql & " (cap_capacita.idnro1 = " & l_ternro & ") AND "
	l_sql = l_sql & " (cap_capacita.origen2 = 3 ) AND "
	l_sql = l_sql & " (cap_capacita.entnro = " & m_rs("modnro") & ") AND "
	l_sql = l_sql & " (cap_capacita.porcen >= " & m_rs("falporcen") & ") "
	rsOpen m_rs2, cn, l_sql, 0
	if not m_rs2.eof then 
		terminada = true
	end if
	m_rs2.Close

	'EVENTO
	l_sql = "SELECT cap_capacita.porcen, cap_capacita.fecha, cap_evento.evecodext "
	l_sql = l_sql & " FROM cap_capacita "
	l_sql = l_sql & " INNER JOIN cap_evento ON ( (cap_capacita.idnro2 = cap_evento.evenro) ) "
	l_sql = l_sql & " WHERE (cap_capacita.origen1 = 4) AND " 'EVENTO
	l_sql = l_sql & " (cap_capacita.idnro1 = " & l_ternro & ") AND "
	l_sql = l_sql & " (cap_capacita.origen2 = 3 ) AND "
	l_sql = l_sql & " (cap_capacita.entnro = " & m_rs("modnro") & ") AND "
	l_sql = l_sql & " (cap_capacita.porcen >= " & m_rs("falporcen") & ") "
	rsOpen m_rs2, cn, l_sql, 0
	if not m_rs2.eof then 
		terminada = true
	end if
	m_rs2.Close

	'CAPACITACION INFORMAL
	l_sql = "SELECT cap_capacita.porcen, cap_estinformal.estinfnro, cap_estinformal.estinffecha, cap_estinformal.estinfdesabr "
	l_sql = l_sql & " FROM cap_capacita "
	l_sql = l_sql & " INNER JOIN cap_estinformal ON cap_capacita.idnro2 = cap_estinformal.estinfnro "
	l_sql = l_sql & " WHERE (cap_capacita.origen1 = 3) AND " 'ESTUDIO INFORMAL
	l_sql = l_sql & " (cap_capacita.idnro1 = " & l_ternro & ") AND "
	l_sql = l_sql & " (cap_capacita.origen2 = 3 ) AND "
	l_sql = l_sql & " (cap_capacita.entnro = " & m_rs("modnro") & ") AND "
	l_sql = l_sql & " (cap_capacita.porcen >= " & m_rs("falporcen") & ") "	
	rsOpen m_rs2, cn, l_sql, 0
	if Not m_rs2.eof then 
		terminada = true
	end if
	m_rs2.Close
	

	'CAPACITACION FORMAL
	l_sql = "SELECT cap_capacita.porcen, titulo.titnro, titulo.titdesabr, cap_estformal.capfecdes "
	l_sql = l_sql & " FROM cap_capacita "
	l_sql = l_sql & " INNER JOIN cap_estformal ON ((cap_capacita.idnro1 = cap_estformal.carredunro) AND (cap_estformal.ternro = " & l_ternro & ")) "
	l_sql = l_sql & " LEFT JOIN titulo ON cap_estformal.titnro = titulo.titnro "
	l_sql = l_sql & " WHERE (cap_capacita.origen1 = 2) AND " 'ESTUDIO FORMAL
	l_sql = l_sql & " (cap_capacita.origen2 = 3 ) AND "
	l_sql = l_sql & " (cap_capacita.entnro = " & m_rs("modnro") & ") AND "
	l_sql = l_sql & " (cap_capacita.porcen >= " & m_rs("falporcen") & ") "		
	rsOpen m_rs2, cn, l_sql, 0
	if Not m_rs2.eof then 
		terminada = true
	end if
	m_rs2.Close


	'ESPECIALIZACIONES
	l_sql = "SELECT cap_capacita.porcen, especemp.espnivnro, especemp.eltananro, espnivel.espnivdesabr, eltoana.eltanadesabr, especemp.espfecha "
	l_sql = l_sql & " FROM cap_capacita "
	l_sql = l_sql & " INNER JOIN especemp ON ( (especemp.ternro = " & l_ternro & ") AND (cap_capacita.idnro1 = especemp.eltananro) AND (cap_capacita.idnro2 = especemp.espnivnro) ) "
	l_sql = l_sql & " INNER JOIN eltoana ON especemp.eltananro = eltoana.eltananro "
	l_sql = l_sql & " INNER JOIN espnivel ON especemp.espnivnro = espnivel.espnivnro "
	l_sql = l_sql & " WHERE (cap_capacita.origen1 = 1) AND " 'ESPECIALIZACION
	l_sql = l_sql & " (cap_capacita.origen2 = 3 ) AND "
	l_sql = l_sql & " (cap_capacita.entnro = " & m_rs("modnro") & ") AND "
	l_sql = l_sql & " (cap_capacita.porcen >= " & m_rs("falporcen") & ") "			
	rsOpen m_rs2, cn, l_sql, 0
	if Not m_rs2.eof then 
		terminada = true
	end if
	m_rs2.Close
		
	' Actualizo la Falencia	
	if terminada = true then 

		l_sql = "UPDATE cap_falencia "
		l_sql = l_sql & "SET falpendiente = 0 " 
		l_sql = l_sql & " WHERE ternro = " & l_ternro 
		l_sql = l_sql & " AND falorigen = 7 "
		l_sql = l_sql & " AND modnro = " & m_rs("modnro")
		m_cm.activeconnection = Cn
		m_cm.CommandText = l_sql
		cmExecute m_cm, l_sql, 0
	else 
		l_sql = "UPDATE cap_falencia "
		l_sql = l_sql & "SET falpendiente = -1 " 
		l_sql = l_sql & " WHERE ternro = " & l_ternro 
		l_sql = l_sql & " AND falorigen = 7 "
		l_sql = l_sql & " AND modnro = " & m_rs("modnro")
		m_cm.activeconnection = Cn
		m_cm.CommandText = l_sql
		cmExecute m_cm, l_sql, 0
	end if 	
	
	m_rs.MoveNext
loop

Set m_cm = Nothing

%>

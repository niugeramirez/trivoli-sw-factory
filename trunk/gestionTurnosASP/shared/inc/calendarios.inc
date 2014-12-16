<%
'Archivo: calendarios.inc
'Descripción: Procedimientos para manejar calendarios
'Autor : Raul Chinestra	
'Fecha: 14/01/2004

 sub FecIniFin(m_evenro)
 
 	Dim m_sql
	m_sql = " SELECT cap_calendario.calfecha FROM cap_eventomodulo " 
	m_sql = m_sql & " INNER JOIN cap_calendario ON cap_calendario.evmonro = cap_eventomodulo.evmonro "		
	m_sql = m_sql & " WHERE cap_eventomodulo.evenro =  " & m_evenro
	m_sql = m_sql & " ORDER BY calfecha "
	rsOpen l_rs, cn, m_sql, 0
	if not (l_rs.eof) then
		  		m_sql = "UPDATE cap_evento "
				m_sql = m_sql & "SET evefecini = " & cambiafecha(l_rs("calfecha"),"YMD",true) 
				m_sql = m_sql & " WHERE evenro = " & m_evenro
				l_cm.activeconnection = Cn
				l_cm.CommandText = m_sql
				cmExecute l_cm, m_sql, 0
	end if

	l_rs.close
	m_sql = " SELECT cap_calendario.calfecha FROM cap_eventomodulo " 
	m_sql = m_sql & " INNER JOIN cap_calendario ON cap_calendario.evmonro = cap_eventomodulo.evmonro "		
	m_sql = m_sql & " WHERE cap_eventomodulo.evenro =  " & m_evenro
	m_sql = m_sql & " ORDER BY calfecha DESC"
	rsOpen l_rs, cn, m_sql, 0
	if not (l_rs.eof) then
		  		m_sql = "UPDATE cap_evento "
				m_sql = m_sql & "SET evefecfin = " & cambiafecha(l_rs("calfecha"),"YMD",true) 
				m_sql = m_sql & " WHERE evenro = " & m_evenro
				l_cm.activeconnection = Cn
				l_cm.CommandText = m_sql
				cmExecute l_cm, m_sql, 0
    end if
	l_rs.close

	
end sub

%>
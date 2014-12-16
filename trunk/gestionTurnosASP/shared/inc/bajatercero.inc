

<%
'Archivo: bajatercero.inc
'Descripción: baja todo lo deterreo
'Autor : CCRossi
'Fecha: 27-03-2004
'Modificado:


 sub bajatercero(l_ternro)
 dim l_sql
 dim l_rs
		 Set l_rs = Server.CreateObject("ADODB.RecordSet")
		'-------------[ BORRO LA RELACION ter_doc ]---------------------
		l_sql = " SELECT tidnro" 
		l_sql = l_sql & " FROM ter_doc "
		l_sql = l_sql & " WHERE ternro = " & l_ternro
		l_rs.Open l_sql, Cn
		do while not l_rs.eof
			l_sql = "DELETE FROM ter_doc WHERE ternro = " & l_ternro
			l_sql = l_sql & "AND tidnro = " & l_rs("tidnro")
			l_cm.activeconnection = Cn
			l_cm.CommandText = l_sql
			cmExecute l_cm, l_sql, 0
			l_rs.MoveNext
		loop
		l_rs.close

		l_sql = " SELECT tipnro" 
		l_sql = l_sql & " FROM ter_tip "
		l_sql = l_sql & " WHERE ternro = " & l_ternro
		l_rs.Open l_sql, Cn
		do while not l_rs.eof
			l_sql = "DELETE FROM ter_tip WHERE ternro = " & l_ternro
			l_sql = l_sql & "AND tipnro = " & l_rs("tipnro")
			l_cm.activeconnection = Cn
			l_cm.CommandText = l_sql
			cmExecute l_cm, l_sql, 0
			l_rs.MoveNext
		loop
		l_rs.close

		' -------------[ BORRO EL DOMICILIO Y LOS TELEFONOS ]---------------------
		l_sql = "SELECT telefono.domnro " &_
		         "FROM cabdom, telefono  " &_
				 "WHERE cabdom.ternro = " & l_ternro & " AND " &_
				 "      telefono.domnro = cabdom.domnro"
		l_rs.Open l_sql, Cn
		do while not l_rs.eof
			l_sql = "DELETE FROM telefono WHERE domnro = " & l_rs("domnro")
			l_cm.activeconnection = Cn
			l_cm.CommandText = l_sql
			cmExecute l_cm, l_sql, 0
			l_rs.movenext
		loop    
		l_rs.close
		
		'----Borro las notas---------
		l_sql = "DELETE FROM notas_ter WHERE ternro = " & l_ternro
		l_cm.activeconnection = Cn
		l_cm.CommandText = l_sql
		cmExecute l_cm, l_sql, 0
		
		
		l_sql = "SELECT detdom.domnro " &_
		         "FROM cabdom, detdom  " &_
				 "WHERE cabdom.ternro = " & l_ternro &_
				 " AND detdom.domnro = cabdom.domnro"
		l_rs.Open l_sql, Cn
		do while not l_rs.eof
				
			l_sql = "DELETE FROM detdom WHERE domnro = " & l_rs("domnro")
			l_cm.activeconnection = Cn
			l_cm.CommandText = l_sql
			cmExecute l_cm, l_sql, 0

'					response.write("<script>alert('borro detdom')</script>")
					
			l_sql = "DELETE FROM cabdom WHERE domnro = " & l_rs("domnro")
			l_cm.activeconnection = Cn
			l_cm.CommandText = l_sql
			cmExecute l_cm, l_sql, 0
				
	'					response.write("<script>alert('borro cab dom')</script>")
			l_rs.movenext
		loop    
		l_rs.close

		' -------------[ BORRO EL TERCERO ]---------------------
		l_sql = "DELETE FROM tercero WHERE ternro = " & l_ternro
		l_cm.activeconnection = Cn
		l_cm.CommandText = l_sql
		response.write ("Tercero: " & l_sql &"<br><br>")
		cmExecute l_cm, l_sql, 0
				
		'		response.write("<script>alert('borro terceroooo')</script>")
				

end sub

%>
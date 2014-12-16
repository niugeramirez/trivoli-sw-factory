<%
'==================================================================
'	include para crear his_estructura para Situacion Revista
'==================================================================

'----------------------------------------------------------------------------------------------------
'Borrar situacion de revista

function borrarSitRevista(ternro, sitRevista, desde, hasta)

	Dim f_rs
	Dim f_cm
	Dim f_sql
	Dim f_desde
	Dim f_hasta
	Dim f_anterior
	Dim f_siguiente
	
	Dim f_sit_ant
	Dim f_sit_sig
	Dim f_sit_sig_hasta
	Dim f_estr_activo
	
	Set f_rs = Server.CreateObject("ADODB.RecordSet")
	Set f_cm = Server.CreateObject("ADODB.Command")
	
	f_desde = cambiafecha(desde,"YMD",true) 
	f_hasta = cambiafecha(hasta,"YMD",true)
	f_anterior   = cambiafecha(DateAdd("d",-1,cDate(desde)),"YMD",true) 
	f_siguiente  = cambiafecha(DateAdd("d", 1,cDate(hasta)),"YMD",true) 
	
	f_sql = "SELECT * FROM his_estructura "
	f_sql = f_sql & " WHERE estrnro = " & sitRevista
	f_sql = f_sql & " AND   tenro   = 30 " 
	f_sql = f_sql & " AND   ternro  = " & ternro
	f_sql = f_sql & " AND   htetdesde  = " & f_desde
	f_sql = f_sql & " AND   htethasta  = " & f_hasta
	
	rsOpen f_rs, cn, f_sql, 0 

	'Me fijo si existe la situacion de revista
	if not f_rs.eof then
		f_rs.Close

		'Borro la situacion de revista
		f_sql = "DELETE FROM his_estructura "
		f_sql = f_sql & " WHERE estrnro = " & sitRevista
		f_sql = f_sql & " AND   tenro   = 30 " 
		f_sql = f_sql & " AND   ternro  = " & ternro
		f_sql = f_sql & " AND   htetdesde  = " & f_desde
		f_sql = f_sql & " AND   htethasta  = " & f_hasta
		
		f_cm.activeconnection = Cn
		f_cm.CommandText = f_sql
		cmExecute f_cm, f_sql, 0
		
		'Busco cual es la estructura de situacion de revista Activo
		f_sql = "SELECT * FROM estructura "
		f_sql = f_sql & " WHERE tenro   = 30 " 
		f_sql = f_sql & " AND   estrcodext = 1 "
		
		rsOpen f_rs, cn, f_sql, 0 
		
		f_estr_activo = -1
		if not f_rs.eof then
			f_estr_activo = f_rs("estrnro")
		end if
		f_rs.close
		
		'Busco el tipo de la situacion de revista anterior
		f_sql = "SELECT * FROM his_estructura "
		f_sql = f_sql & " WHERE tenro   = 30 " 
		f_sql = f_sql & " AND   ternro  = " & ternro
		f_sql = f_sql & " AND   htethasta  = " & f_anterior
		
		rsOpen f_rs, cn, f_sql, 0
		
		f_sit_ant = 0
		if not f_rs.eof then
			f_sit_ant = f_rs("estrnro")
		end if
		f_rs.close
		
		'Busco el tipo de la situacion de revista siguiente
		f_sql = "SELECT * FROM his_estructura "
		f_sql = f_sql & " WHERE tenro   = 30 " 
		f_sql = f_sql & " AND   ternro  = " & ternro
		f_sql = f_sql & " AND   htetdesde  = " & f_siguiente
		
		rsOpen f_rs, cn, f_sql, 0
		
		f_sit_sig = 0
		if not f_rs.eof then
			f_sit_sig = f_rs("estrnro")
			f_sit_sig_hasta = f_rs("htethasta")
		end if
		f_rs.close
		
		'Re-organizamos la situacion de revista de acuerdo al caso
		'Los casos son 4 y son disjuntos

		'+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
		'Caso 1: la sit. de rev. anterior y siguiente son activas
		if CInt(f_estr_activo) = CInt(f_sit_ant) AND CInt(f_estr_activo) = CInt(f_sit_sig) then
			'Borro la situacion de revista siguiente
			f_sql = "DELETE FROM his_estructura "
			f_sql = f_sql & " WHERE tenro   = 30 " 
			f_sql = f_sql & " AND   ternro  = " & ternro
			f_sql = f_sql & " AND   htetdesde  = " & f_siguiente

			f_cm.activeconnection = Cn
			f_cm.CommandText = f_sql
			cmExecute f_cm, f_sql, 0

			if isNull(f_sit_sig_hasta) then
				
				'Modifico la situacion de revista anterior
				f_sql = "UPDATE his_estructura SET "
				f_sql = f_sql & "       htethasta = null "
				f_sql = f_sql & " WHERE tenro   = 30 " 
				f_sql = f_sql & " AND   ternro  = " & ternro
				f_sql = f_sql & " AND   htethasta  = " & f_anterior

			else

				'Modifico la situacion de revista anterior
				f_sql = "UPDATE his_estructura SET "
				f_sql = f_sql & "       htethasta = " & cambiafecha(f_sit_sig_hasta,"YMD",true)
				f_sql = f_sql & " WHERE tenro   = 30 " 
				f_sql = f_sql & " AND   ternro  = " & ternro
				f_sql = f_sql & " AND   htethasta  = " & f_anterior

			end if

			f_cm.activeconnection = Cn
			f_cm.CommandText = f_sql
			cmExecute f_cm, f_sql, 0

		end if
		
		'+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
		'Caso 2: la sit. de rev. anterior es activa y la siguiente no
		if CInt(f_estr_activo) = CInt(f_sit_ant) AND CInt(f_estr_activo) <> CInt(f_sit_sig) then
		
			'Modifico la situacion de revista anterior
			f_sql = "UPDATE his_estructura SET "
			f_sql = f_sql & "       htethasta = " & cambiafecha(hasta,"YMD",true)
			f_sql = f_sql & " WHERE tenro   = 30 " 
			f_sql = f_sql & " AND   ternro  = " & ternro
			f_sql = f_sql & " AND   htethasta  = " & f_anterior

			f_cm.activeconnection = Cn
			f_cm.CommandText = f_sql
			cmExecute f_cm, f_sql, 0

		end if
		
		'+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
		'Caso 3: la sit. de rev. anterior no es activa y la siguiente si
		if CInt(f_estr_activo) <> CInt(f_sit_ant) AND CInt(f_estr_activo) = CInt(f_sit_sig) then
		
			'Modifico la situacion de revista siguiente
			f_sql = "UPDATE his_estructura SET "
			f_sql = f_sql & "       htetdesde = " & cambiafecha(desde,"YMD",true)
			f_sql = f_sql & " WHERE tenro   = 30 " 
			f_sql = f_sql & " AND   ternro  = " & ternro
			f_sql = f_sql & " AND   htetdesde  = " & f_siguiente

			f_cm.activeconnection = Cn
			f_cm.CommandText = f_sql
			cmExecute f_cm, f_sql, 0

		end if
		
		'+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
		'Caso 4: la sit. de rev. anterior y siguiente no son activas
		if CInt(f_estr_activo) <> CInt(f_sit_ant) AND CInt(f_estr_activo) <> CInt(f_sit_sig) then

			'Modifico la situacion de revista siguiente
			f_sql = "INSERT INTO his_estructura "
			f_sql = f_sql & " (tenro, ternro, estrnro, htetdesde,htethasta) "
			f_sql = f_sql & " VALUES (30, " & ternro & ", " 
			f_sql = f_sql & "" & f_estr_activo & ", " 
			f_sql = f_sql & f_desde & ", " 
			f_sql = f_sql & f_hasta & ")"

			f_cm.activeconnection = Cn
			f_cm.CommandText = f_sql
			cmExecute f_cm, f_sql, 0

		end if

	else
	   f_rs.Close
	end if

end function 'borrarSitRevista(ternro, sitRevista, desde, hasta)

'----------------------------------------------------------------------------------------------------
function insertarSitRevista(ternro, sitRevista, desde, hasta)

	Dim f_rs
	Dim f_cm
	Dim f_sql
	Dim f_desde
	Dim f_hasta
	Dim f_primera
	Dim f_ultima
	Dim f_primera_sql
	Dim f_ultima_sql
	Dim f_desde_ant
	Dim f_hasta_sig
	
	Dim f_estr_activo
	
	Set f_rs = Server.CreateObject("ADODB.RecordSet")
	Set f_cm = Server.CreateObject("ADODB.Command")
	
	f_desde = cambiafecha(desde,"YMD",true) 
	f_hasta = cambiafecha(hasta,"YMD",true)
	f_desde_ant  = cambiafecha(DateAdd("d",-1,cDate(desde)),"YMD",true) 
	f_hasta_sig  = cambiafecha(DateAdd("d", 1,cDate(hasta)),"YMD",true) 
	
	f_sql = "SELECT * FROM his_estructura "
	f_sql = f_sql & " WHERE tenro   = 30 " 
	f_sql = f_sql & " AND   ternro  = " & ternro
	f_sql = f_sql & " AND ((  htetdesde <= " & f_desde
	f_sql = f_sql & "     AND htethasta >= " & f_hasta 
	f_sql = f_sql & ") OR ("	
	f_sql = f_sql & "         htetdesde <= " & f_desde
	f_sql = f_sql & "     AND htethasta IS NULL ))"
	
	rsOpen f_rs, cn, f_sql, 0 

	'Me fijo si existe la situacion de revista
	if not f_rs.eof then
		f_primera = f_rs("htetdesde")
		f_ultima  = f_rs("htethasta")

	    f_primera_sql = cambiafecha(f_primera,"YMD",true) 
		if isNull(f_ultima) then
           f_ultima_sql = "NULL"
		else
           f_ultima_sql = cambiafecha(f_ultima,"YMD",true)		
		end if

		f_rs.close
		
		'Busco cual es la estructura de situacion de revista Activo
		f_sql = "SELECT * FROM estructura "
		f_sql = f_sql & " WHERE tenro   = 30 " 
		f_sql = f_sql & " AND   estrcodext = 1 "
		
		rsOpen f_rs, cn, f_sql, 0 
		
		f_estr_activo = -1
		if not f_rs.eof then
			f_estr_activo = f_rs("estrnro")
		end if
		f_rs.close
		
		'Borro la situacion de revista actual
		f_sql = "DELETE FROM his_estructura "
		f_sql = f_sql & " WHERE tenro   = 30 " 
		f_sql = f_sql & " AND   ternro  = " & ternro
		f_sql = f_sql & " AND   htetdesde  = " & f_primera_sql
		if isNull(f_ultima) then
		   f_sql = f_sql & " AND   htethasta IS NULL "
		else
		   f_sql = f_sql & " AND   htethasta  = " & f_ultima_sql
		end if
		
		f_cm.activeconnection = Cn
		f_cm.CommandText = f_sql
		cmExecute f_cm, f_sql, 0

		'Inserto la situacion de revista nueva
		f_sql = "INSERT INTO his_estructura "
		f_sql = f_sql & " (tenro, ternro, estrnro, htetdesde,htethasta) "
		f_sql = f_sql & " VALUES (30, " & ternro & ", " 
		f_sql = f_sql & "" & sitRevista & ", " 
		f_sql = f_sql & f_desde & ", " 
		f_sql = f_sql & f_hasta & ")"

		f_cm.activeconnection = Cn
		f_cm.CommandText = f_sql
		cmExecute f_cm, f_sql, 0
		
		if DateDiff("d",CDate(f_primera),CDate(desde)) <> 0 then
			'Inserto la situacion de revista activo anterior
			f_sql = "INSERT INTO his_estructura "
			f_sql = f_sql & " (tenro, ternro, estrnro, htetdesde,htethasta) "
			f_sql = f_sql & " VALUES (30, " & ternro & ", " 
			f_sql = f_sql & "" & f_estr_activo & ", " 
			f_sql = f_sql & f_primera_sql & ", " 
			f_sql = f_sql & f_desde_ant & ")"
	
			f_cm.activeconnection = Cn
			f_cm.CommandText = f_sql
			cmExecute f_cm, f_sql, 0
		
		end if

		if isNull(f_ultima) then
				'Inserto la situacion de revista activo siguiente
				f_sql = "INSERT INTO his_estructura "
				f_sql = f_sql & " (tenro, ternro, estrnro, htetdesde,htethasta) "
				f_sql = f_sql & " VALUES (30, " & ternro & ", " 
				f_sql = f_sql & "" & f_estr_activo & ", " 
				f_sql = f_sql & f_hasta_sig & ", NULL) " 

				f_cm.activeconnection = Cn
				f_cm.CommandText = f_sql
				cmExecute f_cm, f_sql, 0

		else
			if DateDiff("d",CDate(f_ultima),CDate(hasta)) <> 0 then
				'Inserto la situacion de revista activo siguiente
				f_sql = "INSERT INTO his_estructura "
				f_sql = f_sql & " (tenro, ternro, estrnro, htetdesde,htethasta) "
				f_sql = f_sql & " VALUES (30, " & ternro & ", " 
				f_sql = f_sql & "" & f_estr_activo & ", " 
				f_sql = f_sql & f_hasta_sig & ", " 
				f_sql = f_sql & f_ultima_sql & ")"

				f_cm.activeconnection = Cn
				f_cm.CommandText = f_sql
				cmExecute f_cm, f_sql, 0

			end if
		end if
		
	else
	    f_rs.close
	end if

'----------------------------------------------------------------------------------------------------
end function 'insertarSitRevista(ternro, sitRevista, desde, hasta)


'----------------------------------------------------------------------------------------------------
'Inserta una situacion de revista si corresponde
sub controlarSituacionRevista

	Dim f_rs
	Dim f_cm
	Dim f_sql
	Dim f_desdeanterior
	Dim	f_hastaanterior
	Dim f_tdnro_ant
	Dim f_estrnro
	Dim f_estrnro_ant

	Dim f_estr_activo
	
	Set f_rs = Server.CreateObject("ADODB.RecordSet")
	Set f_cm = Server.CreateObject("ADODB.Command")
	
	'Si es una modificacion borro la sit. de revista anterior
	if (request.queryString("tipo") ="M") then
	
	    'Busca los datos de la licencia anterior
		f_sql = "SELECT * FROM emp_lic "
		f_sql = f_sql & " where emp_licnro = " & l_emp_licnro
		rsOpen f_rs, cn, f_sql, 0 

		if not f_rs.eof then
			f_desdeanterior = f_rs("elfechadesde")
			f_hastaanterior = f_rs("elfechahasta")
			f_tdnro_ant     = f_rs("tdnro")
		end if
		f_rs.close
		
		'Busco la sit. de revista del tipo de dia anterior
		f_sql = "SELECT estrnro, tdnro "
		f_sql = f_sql & "FROM csijp_srtd "
		f_sql = f_sql & "WHERE tdnro = " & f_tdnro_ant
		rsOpen f_rs, cn, f_sql, 0 
		
		f_estrnro_ant = 0
		if not f_rs.eof then
			f_estrnro_ant = f_rs("estrnro") 
		end if	
		f_rs.Close
		
		'Si tiene asociado una situacion de revista borro los datos
		if CInt(f_estrnro_ant) <> 0 then
           borrarSitRevista l_empleado, f_estrnro_ant, f_desdeanterior, f_hastaanterior
		end if

	end if

	'Inserto los datos de la situacion de revista nueva
	
	'Busco la sit. de revista del tipo de dia actual
	f_sql = "SELECT estrnro, tdnro "
	f_sql = f_sql & "FROM csijp_srtd "
	f_sql = f_sql & "WHERE tdnro = " & l_tdnro
	rsOpen f_rs, cn, f_sql, 0 
	
	f_estrnro = 0
	if not f_rs.eof then
		f_estrnro = f_rs("estrnro") 
	end if	
	f_rs.Close
	
	'Si tiene asociado una situacion de revista inserto los datos
	if CInt(f_estrnro) <> 0 then
        insertarSitRevista l_empleado, f_estrnro, l_elfechadesde, l_elfechahasta
	end if

end sub 'controlarSituacionRevista



'----------------------------------------------------------------------------------------------------
'Da de baja una situacion de revista
sub bajaSituacionRevista

	Dim f_rs
	Dim f_sql
	Dim f_desde
	Dim	f_hasta
	Dim f_tdnro
	Dim f_estrnro
	Dim	f_empleado

	Dim f_estr_activo
	
	Set f_rs = Server.CreateObject("ADODB.RecordSet")
	
    'Busca los datos de la licencia
	f_sql = "SELECT * FROM emp_lic "
	f_sql = f_sql & " where emp_licnro = " & l_emp_licnro
	rsOpen f_rs, cn, f_sql, 0 

	if not f_rs.eof then
		f_desde    = f_rs("elfechadesde")
		f_hasta    = f_rs("elfechahasta")
		f_tdnro    = f_rs("tdnro")
		f_empleado = f_rs("empleado")
	end if
	f_rs.close
	
	'Busco la sit. de revista del tipo de dia anterior
	f_sql = "SELECT estrnro, tdnro "
	f_sql = f_sql & "FROM csijp_srtd "
	f_sql = f_sql & "WHERE tdnro = " & f_tdnro
	rsOpen f_rs, cn, f_sql, 0 
	
	f_estrnro = 0
	if not f_rs.eof then
		f_estrnro = f_rs("estrnro") 
	end if	
	f_rs.Close
		
	'Si tiene asociado una situacion de revista borro los datos
	if CInt(f_estrnro) <> 0 then
          borrarSitRevista f_empleado, f_estrnro, f_desde, f_hasta
	end if

end sub 'bajaSituacionRevista

	%>
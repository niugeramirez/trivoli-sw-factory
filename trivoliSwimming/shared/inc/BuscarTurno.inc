<%  on error goto 0
dim NroTurno
dim NombreTurno
dim TipoTurno
dim Tiene_Turno
dim P_asignacion
dim NombreFPago 
dim Nro_FPgo
dim Fecha_inicio
dim P_turcomp 
dim NombreGrupo
dim NroGrupo

NroTurno = 0
Sub BuscarTurno(l_ternro, l_fecha)

Dim salir

  salir = false
    
  ' Si no tiene justificaci¢n, busco los partes de Asignaci¢n de horas
  l_Sql = "SELECT * FROM gti_detturtemp WHERE (ternro = " & l_ternro  & ") AND " 
  l_sql = l_sql & "(gttempdesde <= " & cambiafecha(l_fecha,"YMD",true)  & ") and (" 
  l_sql = l_sql & cambiafecha(l_fecha,"YMD",true)  & " <= gttemphasta)"
  if l_rs.state = 1 then l_rs.close
  l_rs.open l_sql,cn
  If Not l_rs.EOF Then P_asignacion = True

'************* Si no tiene justificación busca los partes de Cambio de Turno
  l_sql = "SELECT gti_turno.turdesabr,gti_turforpago.turnro,gti_turforpago.fpgonro,gti_reldtur.grtddesde, "
  l_sql = l_sql & "gti_reldtur.grtoffset, gti_turno.turcompensa, gti_turno.tipoturno,"   
  l_sql = l_sql & " gti_formapago.fpgodesabr "
  l_sql = l_sql & " FROM  gti_reldtur "
  l_sql = l_sql & " INNER JOIN gti_turforpago ON "
  l_sql = l_sql & " gti_reldtur.turnro = gti_turforpago.turfpagnro "
  l_sql = l_sql & " INNER JOIN gti_turno ON gti_turno.turnro=gti_turforpago.turnro "
  l_sql = l_sql & " INNER JOIN gti_formapago ON gti_turforpago.fpgonro = gti_formapago.fpgonro "
  l_sql = l_sql & " WHERE "
  l_sql = l_sql & " (ternro = "& l_ternro &" ) AND "
  l_sql = l_sql & " (grtddesde <= "& cambiafecha(l_fecha,"YMD",true) &")"
  l_sql = l_sql & " AND (("& cambiafecha(l_fecha,"YMD",true) &" <= grtdhasta) "
  l_sql = l_sql & " OR (grtdhasta is null) ) "

  if l_rs.state = 1 then l_rs.close
  l_rs.open l_sql,cn
  If Not l_rs.EOF Then
      Tiene_Turno = True
      NroTurno = l_rs("turnro")
      NombreTurno = l_rs("turdesabr")
      NombreFPago = l_rs("fpgodesabr")
      Nro_fpgo = l_rs("fpgonro")
      Fecha_inicio = DateAdd("d", l_rs("grtddesde"), -( l_rs("grtoffset")))     
      P_turcomp = l_rs("turcompensa")     'Fecha de inicio del turno
      TipoTurno = l_rs("TipoTurno")
      
      l_sql = " SELECT * FROM his_estructura "
      l_sql = l_sql & " INNER JOIN Alcance_Testr ON his_estructura.tenro = Alcance_Testr.tenro "
      l_sql = l_sql & " INNER JOIN estructura ON his_estructura.estrnro = estructura.estrnro "
      l_sql = l_sql & " WHERE (tanro = " & lngAlcanGrupo & ") AND (ternro = " & l_ternro & ") AND "
      l_sql = l_sql & " (htetdesde <= " & cambiafecha(l_fecha,"YMD",true) & ") AND "
      l_sql = l_sql & " ((" & cambiafecha(l_fecha,"YMD",true) & " <= htethasta) or (htethasta is null))"
      l_sql = l_sql & " ORDER BY alcance_testr.alteorden DESC, his_estructura.htetdesde Desc "
      

      if l_rs.state = 1 then l_rs.close
      l_rs.open l_sql,cn
        If Not l_rs.EOF Then
            NombreGrupo = l_rs("estrdabr")
       	    NroGrupo = l_rs("estrnro")
        End If

      salir = true
  End If

if not salir then

'**** Buscar si la fecha tiene un Turno Asociado en forma Directa en el Histórico		

   l_sql = "SELECT estructura.estrdabr, his_estructura.htetdesde, gti_turfpgogru.*," 
   l_sql = l_sql & " gti_formapago.fpgodesabr, gti_formapago.fpgonro, gti_turno.turnro, "
   l_sql = l_sql & " gti_turno.TipoTurno, gti_turno.turcompensa, gti_turno.turdesabr,alcance_testr.alteorden "
   l_sql = l_sql & " FROM his_estructura "
   l_sql = l_sql & " INNER JOIN Alcance_Testr ON his_estructura.tenro = Alcance_Testr.tenro "
   l_sql = l_sql & " INNER JOIN estructura ON his_estructura.estrnro = estructura.estrnro "
   l_sql = l_sql & " INNER JOIN gti_turfpgogru ON gti_turfpgogru.estrnro = estructura.estrnro "
   l_sql = l_sql & " INNER JOIN gti_turforpago ON "
   l_sql = l_sql & " gti_turforpago.turfpagnro = gti_turfpgogru.turfpagnro " 
   l_sql = l_sql & " INNER JOIN gti_formapago ON gti_formapago.fpgonro = gti_turforpago.fpgonro " 
   l_sql = l_sql & " INNER JOIN gti_turno ON gti_turno.turnro = gti_turforpago.turnro "
   l_sql = l_sql & " WHERE "
   l_sql = l_sql & " (Alcance_Testr.tanro = "&  lngAlcanGrupo &") AND "
   l_sql = l_sql & " (his_estructura.ternro ="&  l_ternro &") AND "
   l_sql = l_sql & " (htetdesde <= "& cambiafecha(l_fecha,"YMD",true) &") AND "
   l_sql = l_sql & " ((htethasta >= "& cambiafecha(l_fecha,"YMD",true) &") "
   l_sql = l_sql & " OR (htethasta is null )) AND "
   l_sql = l_sql & " (fechavalidez <="&  cambiafecha(l_fecha,"YMD",true) & ") "
   l_sql = l_sql & " ORDER BY "
   l_sql = l_sql & " alcance_testr.alteorden DESC, his_estructura.htetdesde DESC, "
   l_sql = l_sql & " gti_turfpgogru.fechavalidez DESC "

	
   if l_rs.state = 1 then l_rs.close
'   l_rs.MaxRecords = 1
   l_rs.open l_sql,cn
		    
   If Not l_rs.EOF Then
	    
      'Existe un turno asociado para la fecha
       Tiene_Turno = True
       NroTurno = l_rs("turnro")
       NombreTurno = l_rs("turdesabr")
       NombreFPago = l_rs("fpgodesabr")
       NombreGrupo = l_rs("estrdabr")
       NroGrupo = l_rs("estrnro")
       TipoTurno = l_rs("TipoTurno")
		        
       P_turcomp = l_rs("turcompensa")
       Fecha_inicio = DateAdd("d", l_rs("fechavalidez"), -(l_rs("offset")))
       Nro_fpgo = l_rs("fpgonro")
   Else

'      '***************** Buscar el Turno Actual del empleado 
'      l_sql = " SELECT estructura.estrdabr,gti_turfpgogru.*,gti_formapago.fpgodesabr, "
'      l_sql = l_sql & " gti_formapago.fpgonro,gti_turno.turnro,gti_turno.TipoTurno, "
'      l_sql = l_sql & " gti_turno.turcompensa,gti_turno.turdesabr,alcance_testr.alteorden "
'      l_sql = l_sql & " FROM estruc_actual "
'      l_sql = l_sql & "	INNER JOIN Alcance_Testr ON estruc_actual.tenro = Alcance_Testr.tenro "
'      l_sql = l_sql & "	INNER JOIN estructura ON estruc_actual.estrnro = estructura.estrnro "
'      l_sql = l_sql & "	INNER JOIN gti_turfpgogru ON gti_turfpgogru.estrnro = estructura.estrnro" 
'      l_sql = l_sql & "	INNER JOIN gti_turforpago ON "
'      l_sql = l_sql & " gti_turforpago.turfpagnro = gti_turfpgogru.turfpagnro "
'      l_sql = l_sql & "	INNER JOIN gti_formapago ON "
'      l_sql = l_sql & " gti_formapago.fpgonro = gti_turforpago.fpgonro "
'      l_sql = l_sql & "	INNER JOIN gti_turno ON gti_turno.turnro = gti_turforpago.turnro "
'      l_sql = l_sql & " WHERE (Alcance_Testr.tanro = "& lngAlcanGrupo &") AND "
'     l_sql = l_sql & " (estruc_actual.ternro = "&  l_ternro &") " 
'      l_sql = l_sql & " ORDER BY "
'      l_sql = l_sql & " alcance_testr.alteorden DESC, gti_turfpgogru.FechaValidez DESC "
'
'		    
'      If l_rs.State <> 0 Then l_rs.Close
'      l_rs.MaxRecords = 1
'      l_rs.open l_sql,cn
'		        
'      If Not l_rs.EOF Then
'            Tiene_Turno = True
'            NroTurno = l_rs("turnro")
'            NombreTurno = l_rs("turdesabr")
'            NombreFPago = l_rs("fpgodesabr")
'            NombreGrupo = l_rs("estrdabr")
'            NroGrupo = l_rs("estrnro")
'            TipoTurno = l_rs("TipoTurno")
'            P_turcomp = l_rs("turcompensa")
'            Fecha_inicio = DateAdd("d", l_rs("fechavalidez"), -l_rs("offset"))
'            Nro_fpgo = l_rs("fpgonro")
'       End If
		        
   End If
		  
end if
		

end sub

%>
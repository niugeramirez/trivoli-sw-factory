<%
'Archivo: concepto.inc
'Descripción: Procedimientos compartidos por los abm de conceptos
'Autor : Alvaro Bayon
'Fecha: 19/11/2003
'Modificado:
'  02-06-04 Alvaro Bayon - Copia de novedades globales

' ===================================================================
' Alcance por ESTRUCTURAS
' ===================================================================
 sub Estructuras(concepto)
	Dim l_seleccionestructuras
	Dim i
	Dim l_lista
	Dim l_tenro
	Dim l_estrnro
	Dim l_sql1
	
	l_sql = "UPDATE concepto "
	l_sql = l_sql & " SET concniv = 1"
	l_sql = l_sql & " WHERE concnro = " & concepto
	l_cm.CommandText = l_sql
	cmExecute l_cm, l_sql, 0
	
	 l_seleccionestructuras = request.Form("seleccionestructuras")
	 l_lista= Split(l_seleccionestructuras,",")		
	 i = 1
	 do while i <= UBound(l_lista)-1
		l_estrnro = l_lista(i)
		l_sql1 = "SELECT tenro FROM estructura "
		l_sql1 = l_sql1 & " WHERE estrnro = " & l_estrnro
		rsOpen l_rs, cn, l_sql1,0
		l_tenro = l_rs("tenro")
		l_rs.close
		
		l_sql = "INSERT INTO cge_segun "
		l_sql = l_sql & "(concnro, nivel, origen, entidad) "
		l_sql = l_sql & " VALUES (" & concepto & ",1," & l_estrnro & "," & l_tenro & ")"
		l_cm.CommandText = l_sql
		cmExecute l_cm, l_sql, 0
		i = i + 1
	 loop
end sub


' ===================================================================
' Alcance GLOBAL
' ===================================================================
  sub global(concepto)
	l_sql = "UPDATE concepto "
	l_sql = l_sql & " SET concniv = 2"
	l_sql = l_sql & " WHERE concnro = " & concepto
	l_cm.CommandText = l_sql
	cmExecute l_cm, l_sql, 0
	
	l_sql = "INSERT INTO cge_segun "
	l_sql = l_sql & "(concnro, nivel, origen, entidad)"
	l_sql = l_sql & " VALUES (" & concepto & ",2,0,0)"
	l_cm.CommandText = l_sql
	cmExecute l_cm, l_sql, 0
  end sub
 
' ===================================================================
' Alcance por EMPLEADO
' ===================================================================
  sub empleado(concepto)
	Dim l_seleccion
	Dim l_ternro
	Dim i
	Dim l_lista
	l_sql = "UPDATE concepto "
	l_sql = l_sql & " SET concniv = 0"
	l_sql = l_sql & " WHERE concnro = " & concepto
	l_cm.CommandText = l_sql
	cmExecute l_cm, l_sql, 0
	 
	 l_seleccion = request.Form("empleados")
	 l_lista= Split(l_seleccion,",")
	 'El primer elemento de la lista es un 0, por lo que comienzo desde 1
	 i = 1
	 do while i <= UBound(l_lista)
		l_ternro = l_lista(i)
		l_sql = "INSERT INTO cge_segun "
		l_sql = l_sql & "(concnro, nivel, origen, entidad) "
		l_sql = l_sql & " VALUES (" & concepto & ",0," & l_ternro & ",0)"
		l_cm.CommandText = l_sql
		cmExecute l_cm, l_sql, 0
		i = i + 1
	 loop
  end sub

sub borrarFormula(concepto)
dim mi_selecc
	l_sql = "SELECT * FROM con_for_tpa "
	l_sql = l_sql & " WHERE concnro = " & concepto
	rsOpenCursor l_rs1, cn, l_sql, 0, adOpenKeyset
	do while not l_rs1.eof
		if isnull(l_rs1("selecc")) then
			mi_selecc = "selecc is null"
		else
			mi_selecc = "selecc='" & l_rs1("selecc") & "'"
		end if
		l_sql = "DELETE con_for_tpa "
		l_sql = l_sql & " WHERE concnro = " & concepto & " AND fornro=" & l_rs1("fornro")
		l_sql = l_sql & " AND tpanro = " & l_rs1("tpanro") & " AND nivel=" & l_rs1("nivel")
		l_sql = l_sql & " AND " & mi_selecc
		l_cm.CommandText = l_sql
		cmExecute l_cm, l_sql, 0
		l_rs1.MoveNext
	loop
	l_rs1.close
end sub

sub insertarFormula()
Dim l_nivel
Dim l_selecc
Dim l_prognro
	'verifico nulos
	if isnull(l_rs("nivel")) or len(l_rs("nivel"))=0 then
		l_nivel = "null"
	else
		l_nivel = l_rs("nivel")
	end if
	if isnull(l_rs("selecc")) or len(l_rs("selecc"))=0 then
		l_selecc = "null"
	else
		l_selecc = "'" & l_rs("selecc") & "'"
	end if
	if isnull(l_rs("prognro")) or len(l_rs("prognro"))=0 then
		l_prognro = "null"
	else
		l_prognro = l_rs("prognro")
	end if
	l_sql = "INSERT into con_for_tpa "
	l_sql = l_sql & "(concnro, tpanro,depurable,cftauto,fornro,nivel,selecc,prognro) "
	l_sql = l_sql & " VALUES (" & l_concnro_nuevo & "," & l_rs("tpanro") 
	l_sql = l_sql & "," & l_rs("depurable") & "," & l_rs("cftauto")
	l_sql = l_sql & "," & l_rs("fornro") & "," & l_nivel
	l_sql = l_sql & "," & l_selecc & "," & l_prognro & ")"
	l_cm.CommandText = l_sql
	cmExecute l_cm, l_sql, 0
end sub

sub borrar_cft_segun(concepto)
dim mi_selecc
dim mi_origen
	l_sql = "SELECT * FROM cft_segun"
	l_sql = l_sql & " WHERE concnro = " & concepto
	rsOpenCursor l_rs1, cn, l_sql, 0, adOpenKeyset
	do while not l_rs1.eof
		if isnull(l_rs1("selecc")) then
			mi_selecc = "selecc is null"
		else
			mi_selecc = "selecc='" & l_rs1("selecc") & "'"
		end if
		if isnull(l_rs1("origen")) then
			mi_origen = "origen is null"
		else
			mi_origen = "origen = " & l_rs1("origen")
		end if
		l_sql = "DELETE cft_segun "
		l_sql = l_sql & " WHERE concnro = " & concepto & " AND fornro=" & l_rs1("fornro")
		l_sql = l_sql & " AND tpanro = " & l_rs1("tpanro") & " AND nivel=" & l_rs1("nivel")
		l_sql = l_sql & " AND " & mi_origen & " AND " & mi_selecc
		l_cm.CommandText = l_sql
		cmExecute l_cm, l_sql, 0
		l_rs1.MoveNext
	loop
	l_rs1.close
end sub

sub insertar_cft_segun()
Dim l_nivel
Dim l_selecc
Dim l_origen
Dim l_entidad
	'verifico nulos
	if isnull(l_rs("nivel")) or len(l_rs("nivel"))=0 then
		l_nivel = "null"
	else
		l_nivel = l_rs("nivel")
	end if
	if isnull(l_rs("selecc")) or len(l_rs("selecc"))=0 then
		l_selecc = "null"
	else
		l_selecc = "'" & l_rs("selecc") & "'"
	end if
	if isnull(l_rs("origen")) or len(l_rs("origen"))=0 then
		l_origen = "null"
	else
		l_origen = l_rs("origen")
	end if
	if isnull(l_rs("entidad")) or len(l_rs("entidad"))=0 then
		l_entidad = "null"
	else
		l_entidad = l_rs("entidad")
	end if
	
	l_sql = "INSERT into cft_segun "
	l_sql = l_sql & "(concnro, tpanro,nivel,origen,selecc,fornro,entidad)"
	l_sql = l_sql & " VALUES(" & l_concnro_nuevo & "," & l_rs("tpanro") 
	l_sql = l_sql & "," & l_nivel & "," & l_origen & "," & l_selecc
	l_sql = l_sql & "," & l_rs("fornro") & "," & l_entidad & ")"
	l_cm.CommandText = l_sql
	cmExecute l_cm, l_sql, 0
end sub

sub borrar_cft_def(concepto)
dim mi_selecc
	l_sql = "SELECT * FROM cft_def"
	l_sql = l_sql & " WHERE concnro = " & concepto
	rsOpenCursor l_rs1, cn, l_sql, 0, adOpenKeyset
	do while not l_rs1.eof
		if isnull(l_rs1("selecc")) then
			mi_selecc = "selecc is null"
		else
			mi_selecc = "selecc='" & l_rs1("selecc") & "'"
		end if
		l_sql = "DELETE cft_def "
		l_sql = l_sql & " WHERE concnro = " & concepto & " AND fornro=" & l_rs1("fornro")
		l_sql = l_sql & " AND tpanro = " & l_rs1("tpanro") & " AND nivelo=" & l_rs1("nivelo")
		l_sql = l_sql & " AND nivelc = " & l_rs1("nivelc") & " AND " & mi_selecc
		l_cm.CommandText = l_sql
		cmExecute l_cm, l_sql, 0
		l_rs1.MoveNext
	loop
	l_rs1.close
end sub

sub insertar_cft_def()
Dim l_nivelc
Dim l_nivelo
Dim l_selecc
Dim l_dftdep
Dim l_dftvalor
	'verifico nulos
	if isnull(l_rs("nivelc")) or len(l_rs("nivelc"))=0 then
		l_nivelc = "null"
	else
		l_nivelc = l_rs("nivelc")
	end if
	if isnull(l_rs("nivelo")) or len(l_rs("nivelo"))=0 then
		l_nivelo = "null"
	else
		l_nivelo = l_rs("nivelo")
	end if
	if isnull(l_rs("selecc")) or len(l_rs("selecc"))=0 then
		l_selecc = "null"
	else
		l_selecc = "'" & l_rs("selecc") & "'"
	end if
	if isnull(l_rs("dftvalor")) or len(l_rs("dftvalor"))=0 then
		l_dftvalor = "null"
	else
		l_dftvalor = replace(l_rs("dftvalor"),",",".")
	end if
	if isnull(l_rs("dftdep")) or len(l_rs("dftdep"))=0 then
		l_dftdep = "null"
	else
		l_dftdep = l_rs("dftdep")
	end if
	
	l_sql = "INSERT into cft_def "
	l_sql = l_sql & "(concnro,tpanro,dftvalor,dftdep,fornro,nivelo,selecc,nivelc)"
	l_sql = l_sql & " VALUES(" & l_concnro_nuevo & "," & l_rs("tpanro") 
	l_sql = l_sql & "," & l_dftvalor & "," & l_rs("dftdep") & "," & l_rs("fornro") 
	l_sql = l_sql & "," & l_nivelo & "," & l_selecc & "," & l_nivelc & ")"
	l_cm.CommandText = l_sql
	cmExecute l_cm, l_sql, 0
end sub
	
sub borrar_cft_resumen(concepto)
	l_sql = "SELECT * FROM cft_resumen"
	l_sql = l_sql & " WHERE concnro = " & concepto
	rsOpenCursor l_rs1, cn, l_sql, 0, adOpenKeyset
	do while not l_rs1.eof
		l_sql = "DELETE cft_resumen "
		l_sql = l_sql & " WHERE concnro = " & concepto & " AND fornro=" & l_rs1("fornro")
		l_sql = l_sql & " AND tpanro = " & l_rs1("tpanro")
		l_cm.CommandText = l_sql
		cmExecute l_cm, l_sql, 0
		l_rs1.MoveNext
	loop
	l_rs1.close
end sub

sub insertar_cft_resumen()
	l_sql = "INSERT into cft_resumen"
	l_sql = l_sql & "(concnro, tpanro,cargl,carest,carind,auto,fornro)"
	l_sql = l_sql & " VALUES(" & l_concnro_nuevo & "," & l_rs("tpanro") 
	l_sql = l_sql & "," & l_rs("cargl") & "," & l_rs("carest") & "," & l_rs("carind") 
	l_sql = l_sql & "," & l_rs("auto") & "," & l_rs("fornro") & ")"
	l_cm.CommandText = l_sql
	cmExecute l_cm, l_sql, 0
end sub

sub borrar_cft_masc(concepto)
	l_sql = "SELECT * FROM cft_masc"
	l_sql = l_sql & " WHERE concnro = " & concepto
	rsOpenCursor l_rs1, cn, l_sql, 0, adOpenKeyset
	do while not l_rs1.eof
		l_sql = "DELETE cft_masc "
		l_sql = l_sql & " WHERE concnro = " & concepto & " AND fornro=" & l_rs1("fornro")
		l_sql = l_sql & " AND tpanro = " & l_rs1("tpanro")
		l_cm.CommandText = l_sql
		cmExecute l_cm, l_sql, 0
		l_rs1.MoveNext
	loop
	l_rs1.close
end sub

sub insertar_cft_masc()
Dim l_parvmin
Dim l_parvmax
Dim l_parmasc
	'verifico nulos
	if isnull(l_rs("parvmin")) or len(l_rs("parvmin"))=0 then
		l_parvmin = "null"
	else
		l_parvmin = replace(l_rs("parvmin"),",",".")
	end if
	if isnull(l_rs("parvmax")) or len(l_rs("parvmax"))=0 then
		l_parvmax = "null"
	else
		l_parvmax = replace(l_rs("parvmax"),",",".")
	end if
	if isnull(l_rs("parmasc")) or len(l_rs("parmasc"))=0 then
		l_parmasc = "null"
	else
		l_parmasc = "'" & l_rs("parmasc") & "'"
	end if
	l_sql = "INSERT into cft_masc "
	l_sql = l_sql & "(concnro, tpanro,parvmin,parvmax,fornro,parmasc) "
	l_sql = l_sql & " VALUES (" & l_concnro_nuevo & "," & l_rs("tpanro") 
	l_sql = l_sql & "," & l_parvmin & "," & l_parvmax
	l_sql = l_sql & "," & l_rs("fornro") & "," & l_parmasc & ")"
	l_cm.CommandText = l_sql
	cmExecute l_cm, l_sql, 0
end sub

sub borrar_novgral(concepto)
Dim l_rs1
	l_sql = "SELECT ngranro FROM novgral"
	l_sql = l_sql & " WHERE concnro = " & concepto
	rsOpenCursor l_rs1, cn, l_sql, 0, adOpenKeyset
	do while not l_rs1.eof
		l_sql = "DELETE novgral "
		l_sql = l_sql & " WHERE ngranro = " & l_rs1("ngranro")
		l_cm.CommandText = l_sql
		cmExecute l_cm, l_sql, 0
		l_rs1.MoveNext
		response.write l_sql
	loop
	l_rs1.close
end sub

sub insertar_novgral()
Dim l_ngravalor
Dim l_ngradesde
Dim l_ngrahasta
Dim l_ngraretro
Dim l_pronro
Dim l_empresa
	'verifico nulos
	if isnull(l_rs("ngravalor")) or len(l_rs("ngravalor"))=0 then
		l_ngravalor = "null"
	else
		l_ngravalor = replace(l_rs("ngravalor"),",",".")
	end if
	if isnull(l_rs("ngradesde")) or len(l_rs("ngradesde"))=0 then
		l_ngradesde = "null"
	else
		l_ngradesde = cambiafecha(l_rs("ngradesde"),"YMD",true)
	end if
	if isnull(l_rs("ngrahasta")) or len(l_rs("ngrahasta"))=0 then
		l_ngrahasta = "null"
	else
		l_ngrahasta = cambiafecha(l_rs("ngrahasta"),"YMD",true)
	end if
	if isnull(l_rs("ngraretro")) or len(l_rs("ngraretro"))=0 then
		l_ngraretro = "null"
	else
		l_ngraretro = cambiafecha(l_rs("ngraretro"),"YMD",true)
	end if
	if isnull(l_rs("pronro")) or len(l_rs("pronro"))=0 then
		l_pronro = "null"
	else
		l_pronro = l_rs("pronro")
	end if
	if isnull(l_rs("empresa")) or len(l_rs("empresa"))=0 then
		l_empresa = "null"
	else
		l_empresa = l_rs("empresa")
	end if

	l_sql = "INSERT into novgral"
	l_sql = l_sql & "(concnro, tpanro,ngravalor,ngravigencia,ngradesde,ngrahasta,ngraretro,pronro,empresa)"
	l_sql = l_sql & " VALUES (" & l_concnro_nuevo & "," & l_rs("tpanro") 
	l_sql = l_sql & "," & l_ngravalor & "," & l_rs("ngravigencia")
	l_sql = l_sql & "," & l_ngradesde & "," & l_ngrahasta & "," & l_ngraretro
	l_sql = l_sql & "," & l_pronro & "," & l_empresa & ")"
	l_cm.CommandText = l_sql
	cmExecute l_cm, l_sql, 0
end sub


sub insertar_paso(concnro,paso)
	l_sql = "SELECT * FROM paso_ext "
	l_sql = l_sql & " WHERE extnro = " & l_concnro & " AND pasnro =" & paso
	rsOpenCursor l_rs, cn, l_sql,0, adOpenKeyset
	if not l_rs.eof then
		l_sql = "INSERT INTO paso_ext"
		l_sql = l_sql & "(extnro, pasnro,extestado,extfecha,extusuario)"
		l_sql = l_sql & " VALUES (" & concnro & "," & paso
		if isnull(l_rs("extestado")) then
			l_sql = l_sql & ",null"
		else
			l_sql = l_sql & "," & l_rs("extestado")
		end if
		l_sql = l_sql & "," & cambiafecha(date(),"YMD",true)
		l_sql = l_sql & ",'" & Session("username") & "'"
		l_sql = l_sql & ")"
		l_cm.CommandText = l_sql
		cmExecute l_cm, l_sql, 0
	end if
	l_rs.close
end sub
	

%>
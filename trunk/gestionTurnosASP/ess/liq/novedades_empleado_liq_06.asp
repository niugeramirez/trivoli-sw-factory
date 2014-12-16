<% Option Explicit %>
<!--#include virtual="/serviciolocal/shared/inc/sec.inc"-->
<!--#include virtual="/serviciolocal/shared/inc/const.inc"-->
<!--#include virtual="/serviciolocal/shared/db/conn_db.inc"-->
<!--#include virtual="/serviciolocal/shared/inc/fecha.inc"-->
<!--
Archivo: novedades_empleado_liq_06.asp
Descripción: abm de novedades del empleado
Autor : FFavre
Fecha: 10/2003
Modificado:
	16-11-03 FFavre Se modifico la secuencia del tab
	19-11-03 FFavre Se modifico la secuencia del tab
	24-11-03 FFavre Se agrego los periodos retroactivos.
 	04-02-04 FFavre No desactiva los campos Periodo desde y hasta al cambiar de un registro a otro.
	12-02-04 FFavre Se verifica que el valor este entre los valores maximos y minimos definidos para el concepto.
    03-09-04 - Scarpa D. - Validacion de los rangos de vigencias	
    06-10-04 - Scarpa D. - Correccion de las novedades retroactivas		
-->
<%
 on error goto 0
 
 Dim l_rs
 Dim l_sql
 
 Dim tipo
 Dim l_empleado
 Dim l_concnro
 Dim l_conccod
 Dim l_concabr
 Dim l_concretro
 Dim l_tpanro
 Dim l_tpadabr
 Dim l_fornro
 Dim l_nevalor
 Dim l_unisigla
 Dim l_conccantdec
 Dim valido
 Dim txt
 Dim l_concretrochec
 Dim l_nepliqddis
 Dim l_nepliqdclass
 Dim l_nepliqdtabIn
 Dim l_nepliqhdis
 Dim l_nepliqhclass
 Dim l_nepliqhtabIn
 Dim l_nepliqdvalue
 Dim l_nepliqhvalue
 
 Dim l_nedesde
 Dim l_nehasta
 Dim l_nenro
 
 l_nedesde  = request.QueryString("nedesde")
 l_nehasta  = request.QueryString("nehasta")
 l_nenro    = request.QueryString("nenro") 

 tipo		= request.QueryString("tipo")
 l_empleado	= request.QueryString("empleado")
 l_concnro 	= request.QueryString("concnro")
 l_conccod 	= request.QueryString("conccod")
 l_tpanro 	= request.QueryString("tpanro")
 l_nevalor  = request.QueryString("nevalor")
 
 if l_nevalor <> "" then
	l_nevalor = l_nevalor / 10000
 end if
 
'=====================================================================================
 function Validar()
 ' Verifica que la novedad no este ya cargada (empleado, concnro, tpanro) en el caso de un alta
' 	if tipo = "M" then
'		Validar = true
'	else
'		l_sql = "SELECT * "
'		l_sql = l_sql & "FROM novemp "
'		l_sql = l_sql & "WHERE empleado = " & l_empleado & " AND concnro = " & l_concnro & " AND tpanro = " & l_tpanro
'		rsOpen l_rs, cn, l_sql, 0
'		if not l_rs.eof then
'			response.write "parent.invalido('duplicada');"
'			Validar = false
'		else
'			Validar = true
'		end if
'		l_rs.Close
'	end if

Validar = true

 end function
'=====================================================================================
 response.write "<script>" & vbCrLf
 
 Set l_rs = Server.CreateObject("ADODB.RecordSet")
 
 select case tipo
	case "concepto"
		l_sql = "SELECT concepto.concnro, concepto.concabr, concepto.concretro, concepto.conccantdec "
		l_sql = l_sql & "FROM concepto "
		l_sql = l_sql & "WHERE conccod = '" &  l_conccod & "'"
		rsOpen l_rs, cn, l_sql, 0
		
		l_concretrochec = "false"
		l_nepliqddis = "true"
		l_nepliqdvalue = "0,,"
		l_nepliqdclass = "deshabinp"
		l_nepliqdtabIn = "-1"
		l_nepliqhvalue = "0,,"
		l_nepliqhdis = "true"
		l_nepliqhclass = "deshabinp"
		l_nepliqhtabIn = "-1"
		
		if l_rs.eof then
			' No existe el concepto
			response.write "parent.invalido('ConcNoExiste');" & vbCrLf
			l_concnro = 0
			l_conccod = ""
			l_concabr = ""
			l_tpanro  = ""
			l_tpadabr = ""
			l_conccantdec = ""
		else
			' Existe el concepto
			l_concnro = l_rs("concnro")
			l_concabr = l_rs("concabr")
			l_concretro = l_rs("concretro")
			l_conccantdec = l_rs("conccantdec")
			l_rs.close
			l_sql = "SELECT cft_resumen.tpanro, cft_resumen.carind "
			l_sql = l_sql & "FROM cft_resumen "
			l_sql = l_sql & "WHERE concnro = " &  l_concnro & " AND carind = -1"
			rsOpen l_rs, cn, l_sql, 0
			if l_rs.eof then
				' No es un concepto individual
				response.write "parent.invalido('NoConceptoInd');" & vbCrLf
				l_tpanro  = ""
				l_tpadabr = ""
			else
				' Es valido.
		 		l_tpanro = l_rs("tpanro")
				l_rs.close
				l_sql = "SELECT tipopar.tpanro, tipopar.tpadabr, unidad.unisigla "
				l_sql = l_sql & "FROM tipopar INNER JOIN unidad ON tipopar.uninro = unidad.uninro "
				l_sql = l_sql & "WHERE tipopar.tpanro = " & l_tpanro
				rsOpen l_rs, cn, l_sql, 0
				l_tpadabr = l_rs("tpadabr")
				l_unisigla = l_rs("unisigla")
				if l_concretro = true then
					l_concretrochec = "true"
					l_nepliqddis = "false"
					l_nepliqdclass = "habinp"
					l_nepliqdtabIn = "6"
					l_nepliqhdis = "false"
					l_nepliqhclass = "habinp"
					l_nepliqhtabIn = "7"
				end if
 				l_rs.close
			end if
		end if
		response.write "parent.document.all.concnro.value='" & l_concnro & "';" & vbCrLf
		response.write "parent.document.all.conccod.value='" & l_conccod & "';" & vbCrLf
		response.write "parent.document.all.concabr.value='" & l_concabr & "';" & vbCrLf
		response.write "parent.document.all.conccantdec.value='" & l_conccantdec & "';" & vbCrLf
		response.write "parent.document.all.tpanro.value='" & l_tpanro & "';" & vbCrLf
		response.write "parent.document.all.tpadabr.value='" & l_tpadabr & "';" & vbCrLf
		
		response.write "parent.actualizarConcRetro(" & l_concretrochec & ");" & vbCrLf
		'response.write "parent.document.all.concretro.checked = " & l_concretrochec & ";" & vbCrLf
		'response.write "parent.document.all.nepliqdesde.value = '" & l_nepliqdvalue & "';" & vbCrLf
		'response.write "parent.document.all.nepliqdesde.disabled = " & l_nepliqddis & ";" & vbCrLf
		'response.write "parent.document.all.nepliqdesde.className = '" & l_nepliqdclass & "';" & vbCrLf
		'response.write "parent.document.all.nepliqdesde.tabIndex = " & l_nepliqdtabIn & ";" & vbCrLf
		'response.write "parent.document.all.nepliqhasta.value = '" & l_nepliqhvalue & "';" & vbCrLf
		'response.write "parent.document.all.nepliqhasta.disabled = " & l_nepliqhdis & ";" & vbCrLf
		'response.write "parent.document.all.nepliqhasta.className = '" & l_nepliqhclass & "';" & vbCrLf
		'response.write "parent.document.all.nepliqhasta.tabIndex = " & l_nepliqhtabIn & ";" & vbCrLf
		
		if l_tpanro <> "" then
			' Se encontro un parametro.
			response.write "parent.document.all.unidesc.value='" & l_unisigla & "';" & vbCrLf
			response.write "parent.document.all.nevalor.focus();" & vbCrLf
			response.write "parent.document.all.nevalor.select();" & vbCrLf
		else
			response.write "parent.document.all.conccod.focus();" & vbCrLf
			response.write "parent.document.all.conccod.select();" & vbCrLf
		end if
	
	
	case "parametro"
		if l_concnro <> "" then
			' Se verifica que se haya cargado un concepto
			l_sql = "SELECT tipopar.tpanro, tipopar.tpadabr, unidad.unisigla "
			l_sql = l_sql & "FROM tipopar INNER JOIN con_for_tpa ON tipopar.tpanro = con_for_tpa.tpanro " 
			l_sql = l_sql & "INNER JOIN unidad ON tipopar.uninro = unidad.uninro "
			l_sql = l_sql & "WHERE con_for_tpa.concnro = " & l_concnro & " AND con_for_tpa.tpanro = " & l_tpanro
			rsOpen l_rs, cn, l_sql, 0
			if l_rs.eof then
				' No esta asignado al concepto
				response.write "parent.invalido('ParNoValido');" & vbCrLf
				response.write "parent.document.all.tpanro.value='';" & vbCrLf
				response.write "parent.document.all.tpanrold.value='';" & vbCrLf
				response.write "parent.document.all.unidesc.value='';" & vbCrLf
				response.write "parent.document.all.tpadabr.value='';" & vbCrLf
				response.write "parent.document.all.tpanro.focus();" & vbCrLf
				response.write "parent.document.all.tpanro.select();" & vbCrLf
			else
				' Esta asignado al concepto
				response.write "parent.document.all.tpanrold.value=" & l_tpanro & ";" & vbCrLf
				response.write "parent.document.all.unidesc.value='" & l_rs("unisigla") & "';" & vbCrLf
				response.write "parent.document.all.tpadabr.value='" & l_rs("tpadabr") & "';" & vbCrLf
				response.write "parent.document.all.nevalor.focus();" & vbCrLf
				response.write "parent.document.all.nevalor.select();" & vbCrLf
			end if
			l_rs.close
		end if
	
	
	' Verifica que los datos sean validos antes de guardarlos
	case "A", "M"
		' Verifica que sea valida la novedad.
		if Validar() then
			l_sql = "SELECT parvmin, parvmax "
			l_sql = l_sql & "FROM cft_masc "
			l_sql = l_sql & "WHERE concnro = " & l_concnro & " AND tpanro = " & l_tpanro
			rsOpen l_rs, cn, l_sql, 0
			valido = true
			' Verifica que el valor este comprendidos entre los valores permitidos
			if not l_rs.eof then
				if (CDbl(l_rs("parvmin")) <> CDbl(0) and CDbl(l_rs("parvmin")) > CDbl(l_nevalor)) then
					txt = "El Valor es menor que el valor mínimo (" & l_rs("parvmin") & ") permitido para el parámetro."
					response.write "alert('" & txt & "');"
					response.write "parent.document.all.nevalor.focus();" & vbCrLf
					response.write "parent.document.all.nevalor.select();" & vbCrLf
					valido = false
				else
					if (CDbl(l_rs("parvmax")) <> CDbl(0) and (CDbl(l_rs("parvmax")) < CDbl(l_nevalor))) then
						txt = "El Valor es mayor que el valor máximo (" & l_rs("parvmax") & ") permitido para el parámetro."
						response.write "alert('" & txt & "');"
						response.write "parent.document.all.nevalor.focus();" & vbCrLf
						response.write "parent.document.all.nevalor.select();" & vbCrLf
						valido = false
					end if
				end if
			end if
			l_rs.close
			if valido then
			    Dim ini
				Dim fin
				
				ini = l_nedesde
				fin = l_nehasta
			
			    if trim(l_nedesde) = "" then

			        'Controlo si existe una novedad igual
					l_sql = "SELECT * "
					l_sql = l_sql & " FROM novemp "
					l_sql = l_sql & " WHERE empleado = " & l_empleado & " AND concnro = " & l_concnro & " AND tpanro = " & l_tpanro

					if tipo = "M" then
					   l_sql = l_sql & " AND nenro <> "	& l_nenro
					end if

				    rsOpen l_rs, cn, l_sql, 0

					if not l_rs.eof then
					   response.write "parent.invalido('duplicada');"
					else
                       response.write "parent.Valido();"
					end if

					l_rs.close

				else
				
				    'Controlo si existe una novedad sin vigencia
					l_sql = "SELECT * "
					l_sql = l_sql & " FROM novemp "
					l_sql = l_sql & " WHERE empleado = " & l_empleado & " AND concnro = " & l_concnro & " AND tpanro = " & l_tpanro
					l_sql = l_sql & " AND nevigencia=0 "
					if tipo = "M" then
					   l_sql = l_sql & " AND nenro <> "	& l_nenro
					end if
					
				    rsOpen l_rs, cn, l_sql, 0
					
					if not l_rs.eof then
					   response.write "parent.invalido('Vigencia01');"
					else

						l_rs.close

						if trim(l_nehasta) = "" then
						   'Controlos si hay alguna con hasta nulo

							l_sql = "SELECT * "
							l_sql = l_sql & " FROM novemp "
							l_sql = l_sql & " WHERE empleado = " & l_empleado & " AND concnro = " & l_concnro & " AND tpanro = " & l_tpanro
							l_sql = l_sql & " AND nevigencia = -1 "
							l_sql = l_sql & " AND nehasta is null "

							if tipo = "M" then
							   l_sql = l_sql & " AND nenro <> "	& l_nenro
							end if

						    rsOpen l_rs, cn, l_sql, 0

							if not l_rs.eof then
							   response.write "parent.invalido('Vigencia03');</script>"
							   response.end							   
							else
						        'Controlo si existe una novedad con fecha mayor a la fecha desde de esta vigencia
						        l_rs.close
								
								l_sql = "SELECT * "
								l_sql = l_sql & " FROM novemp "
								l_sql = l_sql & " WHERE empleado = " & l_empleado & " AND concnro = " & l_concnro & " AND tpanro = " & l_tpanro
								l_sql = l_sql & " AND nevigencia = -1 "
								l_sql = l_sql & " AND (nedesde >= " & cambiafecha(ini,"YMD",true)
								l_sql = l_sql & "  OR  nehasta >= " & cambiafecha(ini,"YMD",true) & ") "

								if tipo = "M" then
								   l_sql = l_sql & " AND nenro <> "	& l_nenro
								end if

							    rsOpen l_rs, cn, l_sql, 0

								if not l_rs.eof then
								   response.write "parent.invalido('Vigencia02');</script>"
								   response.end
								else
								   response.write "parent.Valido();</script>"
   								   response.end
								end if

						    end if

						else
						
					        'Controlo si existe una novedad con fecha hasta nulo y desde menor o igual al desde
							l_sql = "SELECT * "
							l_sql = l_sql & " FROM novemp "
							l_sql = l_sql & " WHERE empleado = " & l_empleado & " AND concnro = " & l_concnro & " AND tpanro = " & l_tpanro
							l_sql = l_sql & " AND nevigencia = -1 "
							l_sql = l_sql & " AND (nedesde <= " & cambiafecha(ini,"YMD",true)
							l_sql = l_sql & "  OR  nedesde <= " & cambiafecha(fin,"YMD",true) & ") "
							l_sql = l_sql & " AND  nehasta IS NULL "

							if tipo = "M" then
							   l_sql = l_sql & " AND nenro <> "	& l_nenro
							end if

						    rsOpen l_rs, cn, l_sql, 0

							if not l_rs.eof then
							   response.write "parent.invalido('Vigencia02');</script>"
							   response.end							   
							else
							    l_rs.close

							    'Controlo si existe una novedad sin vigencia
								l_sql = "SELECT * "
								l_sql = l_sql & " FROM novemp "
								l_sql = l_sql & " WHERE empleado = " & l_empleado & " AND concnro = " & l_concnro & " AND tpanro = " & l_tpanro
								l_sql = l_sql & " AND nevigencia = -1 "
								l_sql = l_sql & " AND ((nedesde >=" & cambiafecha(ini,"YMD",true)
								l_sql = l_sql & " and nehasta <= " & cambiafecha(fin,"YMD",true) & ") "
								l_sql = l_sql & " or (nedesde <  " & cambiafecha(ini,"YMD",true)
								l_sql = l_sql & " and nehasta <= " & cambiafecha(fin,"YMD",true) 
								l_sql = l_sql & " and nehasta >= " & cambiafecha(ini,"YMD",true) & ") "	
								l_sql = l_sql & " or (nedesde >= " & cambiafecha(ini,"YMD",true)
								l_sql = l_sql & " and nehasta >  " & cambiafecha(fin,"YMD",true) 
								l_sql = l_sql & " and nedesde <= " & cambiafecha(fin,"YMD",true) & ") "	
								l_sql = l_sql & " or (nedesde <  " & cambiafecha(ini,"YMD",true)
								l_sql = l_sql & " and nehasta >  " & cambiafecha(fin,"YMD",true) & ")) "

								if tipo = "M" then
								   l_sql = l_sql & " AND nenro <> "	& l_nenro
								end if

							    rsOpen l_rs, cn, l_sql, 0
								
								if not l_rs.eof then
								   response.write "parent.invalido('Vigencia02');"
								else
								   response.write "parent.Valido();</script>"
								   response.end
								end if

							end if

						end if

					end if
					
					l_rs.close

				end if

			end if
		end if
 end select
 
 response.write "</script>" & vbCrLf
 Set l_rs = Nothing
%>

<% Option Explicit %>
<!--#include virtual="/rhprox2/shared/inc/sec.inc"-->
<!--#include virtual="/rhprox2/shared/inc/const.inc"-->
<!--#include virtual="/rhprox2/shared/db/conn_db.inc"-->
<!--#include virtual="/rhprox2/shared/inc/fecha.inc"-->
<!--#include virtual="/rhprox2/shared/inc/calendarios.inc"-->
<!--#include virtual="/rhprox2/shared/inc/asistente.inc"-->
<!--#include virtual="/rhprox2/shared/inc/adovbs.inc"-->
<!--
Archivo: calendarios_cap_08.asp
Descripción: Creacion de Calendarios en Forma Masiva
Autor : Raul CHinestra
Fecha: 25/10/2003
Midificado : xx/02/2004 - Lisandro Moro   - Superposicion de calendarios
Midificado : 26/02/2004 - Lisandro Moro   - Superposicion de calendarios con los de distintos modulos..
Midificado : 15/03/2004 - Muzzolón Martín - Se agregaron los días sabado y domingo..
-->

<script src="/rhprox2/shared/js/fn_calendarios.vbs" language="VBScript" runat="Server"></script>

<% 
on error goto 0

Dim l_tipo
Dim l_rs
Dim l_rs2
Dim l_cm
Dim l_sql
Dim l_sql2
Dim l_evmonro
Dim l_evento
Dim l_evenro

l_evmonro       = request.querystring("evmo")

Dim l_lugnro
Dim l_caldia
Dim l_calfecini
Dim l_calfecfin
Dim l_calhordes1
Dim l_calhordes2
Dim l_calhorhas1
Dim l_calhorhas2
Dim l_rbopc
Dim l_lu
Dim l_ma
Dim l_mi
Dim l_ju
Dim l_vi
Dim l_sa
Dim l_do

Dim l_fec
Dim l_fecha
Dim l_fini
Dim l_ffin
Dim l_contador
Dim l_aux
Dim l_semana
Dim l_fecpar

Dim l_fecha_temp
Dim l_superpuesto
Dim l_semana_incorrecta

l_lugnro	 	= request.Form("lugnro")
l_calfecini	 	= request.Form("calfecini")
l_calfecfin	 	= request.Form("calfecfin")
l_calhordes1    = request.Form("calhordes1")
l_calhordes2    = request.Form("calhordes2")
l_calhorhas1    = request.Form("calhorhas1")
l_calhorhas2    = request.Form("calhorhas2")
l_rbopc         = request.Form("rbopc")
l_lu            = request.Form("lu")
l_ma            = request.Form("ma")
l_mi            = request.Form("mi")
l_ju            = request.Form("ju")
l_vi            = request.Form("vi")
l_sa            = request.Form("sa")
l_do            = request.Form("dom")

l_semana        = request.Form("semana")

l_fecha = CDate(l_calfecini)

l_semana_incorrecta = false

function haysuperposicion

haysuperposicion = false

Dim f_rs
Dim f_rs2

Set f_rs = Server.CreateObject("ADODB.RecordSet")
Set f_rs2 = Server.CreateObject("ADODB.RecordSet")

' Controlo que no se superponga en horario y lugar para calendarios del mismo modulo.
l_sql = "SELECT calnro "
l_sql = l_sql & " FROM cap_calendario "
l_sql = l_sql & " WHERE calfecha=" & cambiafecha(l_fecha_temp,"YMD",true)
l_sql = l_sql & " and evmonro=" & l_evmonro
l_sql = l_sql & " and (('" & l_calhordes1 & l_calhordes2 & "' < calhordes and '" & l_calhorhas1 & l_calhorhas2 & "' >= calhordes) "
l_sql = l_sql & " or ('" & l_calhordes1 & l_calhordes2 & "'>= calhordes and '" & l_calhordes1 & l_calhordes2 & "' <= calhorhas)) "
l_sql = l_sql & " and lugnro = " & l_lugnro
rsOpen f_rs, cn, l_sql, 0
if not f_rs.eof then
	f_rs.close
    haysuperposicion = true
    set f_rs = nothing
else
	'Controlo que no se superponga con el calendario de otros modulos de este evento. *****'
	f_rs.close
	l_sql = " SELECT evenro "
	l_sql = l_sql & " FROM cap_eventomodulo "
	l_sql = l_sql & " WHERE evmonro = " & l_evmonro
	rsOpen f_rs, cn, l_sql, 0	
	l_evenro = f_rs("evenro")
	f_rs.Close
	l_sql = " SELECT evmonro "
	l_sql = l_sql & " FROM cap_eventomodulo "
	l_sql = l_sql & " WHERE evenro = " & l_evenro
	l_sql = l_sql & " AND evmonro <> " & l_evmonro
	rsOpen f_rs, cn, l_sql, 0
	do until f_rs.eof
		l_sql = "SELECT calnro "
		l_sql = l_sql & " FROM cap_calendario "
		l_sql = l_sql & " WHERE calfecha =" & cambiafecha(l_fecha_temp,"YMD",true)
		l_sql = l_sql & " and evmonro =" & f_rs("evmonro")
		l_sql = l_sql & " and (('" & l_calhordes1 & l_calhordes2 & "' < calhordes and '" & l_calhorhas1 & l_calhorhas2 & "' >= calhordes) "
		l_sql = l_sql & " or ('" & l_calhordes1 & l_calhordes2 & "'>= calhordes and '" & l_calhordes1 & l_calhordes2 & "' <= calhorhas)) "
		'response.write l_sql & "<br>"
		rsOpen f_rs2, cn, l_sql, 0
		if not f_rs2.eof then
			f_rs2.close
		    haysuperposicion = true
		else
			f_rs2.close
		end if
		f_rs.MoveNext
	loop
	
	if haysuperposicion = false then 

		'Controlo que no se superponga con el calendario de otros modulos de este evento. *****'
		l_sql = "SELECT calnro "
		l_sql = l_sql & " FROM cap_calendario "
		l_sql = l_sql & " WHERE calfecha=" & cambiafecha(l_fecha_temp,"YMD",true)
		l_sql = l_sql & " and (('" & l_calhordes1 & l_calhordes2 & "' < calhordes and '" & l_calhorhas1 & l_calhorhas2 & "' >= calhordes) "
		l_sql = l_sql & " or ('" & l_calhordes1 & l_calhordes2 & "'>= calhordes and '" & l_calhordes1 & l_calhordes2 & "' <= calhorhas)) "
		l_sql = l_sql & " and lugnro = " & l_lugnro
		rsOpen f_rs2, cn, l_sql, 0
		if not f_rs2.eof then		
		    haysuperposicion = true
		else
			haysuperposicion = false
		end if
		f_rs2.close
	end if
end if


Set f_rs = nothing
Set f_rs2 = nothing

end function


'Creo los objetos ADO
set l_rs = Server.CreateObject("ADODB.RecordSet")
set l_cm = Server.CreateObject("ADODB.Command")
set l_rs2 = Server.CreateObject("ADODB.RecordSet")

'---------------------------------------------------
l_fecha_temp = l_fecha
l_superpuesto = 0
	
l_sql2 = " SELECT evenro "
l_sql2 = l_sql2 & " FROM cap_eventomodulo "
l_sql2 = l_sql2 & " WHERE evmonro = " & l_evmonro
rsOpen l_rs2, cn, l_sql2, 0
if not l_rs2.eof then 'una ves que obtengo el nro de evento, busco los numeros de eventos-modulos para ese evento
	l_evenro = l_rs2("evenro")
	l_rs2.Close
	l_sql2 = " SELECT evmonro "
	l_sql2 = l_sql2 & " FROM cap_eventomodulo "
	l_sql2 = l_sql2 & " WHERE evenro = " & l_evenro
	rsOpen l_rs2, cn, l_sql2, 0
	l_rs2.MoveFirst
	do until l_rs2.eof
			l_fecha_temp = l_fecha
			Select Case l_rbopc
			           Case "1"  
			                Do While DateDiff("d", l_fecha_temp, CDate(l_calfecfin)) >= 0 and l_superpuesto <> 1
								if haysuperposicion = true then 
									l_superpuesto = 1
								end if
			                    l_fecha_temp = DateAdd("d", 1, l_fecha_temp)
			                Loop
			           Case "2"  
			           		Do While DateDiff("d", l_fecha_temp, CDate(l_calfecfin)) >= 0
						 		if (l_lu = "on" and weekday(l_fecha_temp) = 2) or (l_ma = "on" and weekday(l_fecha_temp) = 3) or (l_mi = "on" and weekday(l_fecha_temp) = 4) or (l_ju = "on" and weekday(l_fecha_temp) = 5) or (l_vi = "on" and weekday(l_fecha_temp) = 6) or (l_sa = "on" and weekday(l_fecha_temp) = 7) or (l_do = "on" and weekday(l_fecha_temp) = 1) then
									if haysuperposicion = true then 
										l_superpuesto = 1
									end if
								end if
								l_fecha_temp = DateAdd("d", 1, l_fecha_temp)
							Loop 
			           Case "3"
					   		l_semana_incorrecta = true  
					   		l_fini = DateValue("01/" & CStr(Month(l_fecha)) & "/" & CStr(year(l_fecha)) )
						    if Month(l_calfecfin) = 12 then
			          			l_ffin = DateValue ("31/" & CStr(Month(l_calfecfin)) & "/" + CSTr(year(l_calfecfin)) )
						    else 
			                    l_ffin = datevalue ("01/" & CStr(month(l_calfecfin) + 1 ) & "/" + CStr(year(l_calfecfin))) - 1
							end if
							While l_fini < l_ffin
								For l_contador = 1 To 7
								   l_aux = DateValue ( CStr(l_contador)  + "/" + CStr(month(l_fini)) + "/" + CStr(year(l_fini)))
								   if l_aux <= l_calfecfin then
				         		 		if (l_lu = "on" and weekday(l_aux) = 2) or (l_ma = "on" and weekday(l_aux) = 3) or (l_mi = "on" and weekday(l_aux) = 4) or (l_ju = "on" and weekday(l_aux) = 5) or (l_vi = "on" and weekday(l_aux) = 6 )  or (l_sa = "on" and weekday(l_aux) = 7) or (l_do = "on" and weekday(l_aux) = 1) then
										   	   l_fecha_temp =  dateadd("d", 7 * (Cint(l_semana) - 1), l_aux)
											   l_caldia = DiadeSemana(l_fecpar)
												'Verifico que un dia de la semana seleccionada caiga dentro del rango fecha desde - hasta:
												if l_fecha_temp <= CDate(l_calfecfin) and l_fecha_temp >= CDate(l_calfecini)  then 
													l_semana_incorrecta = false
												end if
												
												if haysuperposicion = true then 
													l_superpuesto = 1
												end if	
									    end if	
									end if	
								Next 
								if month(l_fini) = 12 then
						            l_fini = datevalue ("01/" & "01" & "/" + CStr(year(l_fini) + 1) )
						        else 
			                        l_fini = datevalue ("01/" &  CStr(month(l_fini) + 1) & "/" + CStr(year(l_fini)))
								end if
							Wend
						'l_rs.close
			end select
		l_rs2.MoveNext
	loop
end if

if l_superpuesto = 1 then
	%><script>alert("Los Calendarios que se desean generar se superponen con los existentes.");window.close();</script><%
	response.end
end if 

if l_semana_incorrecta = true then
	%><script>alert("Los datos ingresados no generarán ningún calendario. Verificar los mismos.");window.close();</script><%
	response.end
end if 



'---------------------------------------------------


Select Case l_rbopc
           Case "1"  
				l_cm.activeconnection = Cn
                    Do While DateDiff("d", l_fecha, CDate(l_calfecfin)) >= 0
    				   l_caldia = DiadeSemana(l_fecha)
    				   l_sql = "INSERT INTO cap_calendario "
    				   l_sql = l_sql & "(calfecha, caldia, calhordes, calhorhas, lugnro, evmonro ) "
    				   l_sql = l_sql & "VALUES (" & cambiafecha(l_fecha,"YMD",true)  & ",'" & l_caldia & "','" & l_calhordes1 & l_calhordes2 & "','" & l_calhorhas1 & l_calhorhas2 & "'," & l_lugnro & "," & l_evmonro & ")"   
    				   l_cm.CommandText = l_sql
    				   cmExecute l_cm, l_sql, 0   
    				   l_fecha = DateAdd("d", 1, l_fecha)
    				Loop
				'l_rs.close
           Case "2"  
    		   	l_cm.activeconnection = Cn
				Do While DateDiff("d", l_fecha, CDate(l_calfecfin)) >= 0
		 		if (l_lu = "on" and weekday(l_fecha) = 2) or _
				   (l_ma = "on" and weekday(l_fecha) = 3) or _
	 			   (l_mi = "on" and weekday(l_fecha) = 4) or _
				   (l_ju = "on" and weekday(l_fecha) = 5) or _
				   (l_vi = "on" and weekday(l_fecha) = 6) or _
				   (l_sa = "on" and weekday(l_fecha) = 7) or _
				   (l_do = "on" and weekday(l_fecha) = 1) then
				   		l_caldia = DiadeSemana(l_fecha)
	       			    l_sql = "INSERT INTO cap_calendario "
				        l_sql = l_sql & "(calfecha, caldia, calhordes, calhorhas, lugnro, evmonro ) "
				        l_sql = l_sql & "VALUES (" & cambiafecha(l_fecha,"YMD",true)  & ",'" & l_caldia & "','" & l_calhordes1 & l_calhordes2 & "','" & l_calhorhas1 & l_calhorhas2 & "'," & l_lugnro & "," & l_evmonro & ")"
   					    l_cm.CommandText = l_sql
				        cmExecute l_cm, l_sql, 0   
				end if		
				l_fecha = DateAdd("d", 1, l_fecha)
				Loop 
				'l_rs.close
				
           Case "3"
				
		   		l_fini = DateValue("01/" & CStr(Month(l_fecha)) & "/" & CStr(year(l_fecha)) )

			    if Month(l_calfecfin) = 12 then
          			l_ffin = DateValue ("31/" & CStr(Month(l_calfecfin)) & "/" + CSTr(year(l_calfecfin)) )
			    else 
                    l_ffin = datevalue ("01/" & CStr(month(l_calfecfin) + 1 ) & "/" + CStr(year(l_calfecfin))) - 1 		   
				end if
				
    		   	l_cm.activeconnection = Cn
				
				While l_fini < l_ffin
				
					For l_contador = 1 To 7
					
					   l_aux = DateValue ( CStr(l_contador)  + "/" + CStr(month(l_fini)) + "/" + CStr(year(l_fini)))
					   
					   if l_aux <= l_calfecfin then             
	         		 		if (l_lu = "on" and weekday(l_aux) = 2) or _
				  			   (l_ma = "on" and weekday(l_aux) = 3) or _
				 			   (l_mi = "on" and weekday(l_aux) = 4) or _
							   (l_ju = "on" and weekday(l_aux) = 5) or _
							   (l_vi = "on" and weekday(l_aux) = 6) or _
   							   (l_sa = "on" and weekday(l_aux) = 7) or _
							   (l_do = "on" and weekday(l_aux) = 1) then
							   	   	l_fecpar =  dateadd("d", 7 * (Cint(l_semana) - 1), l_aux)
								   	l_caldia = DiadeSemana(l_fecpar)
								   	'verifico que el dia este dentro del fecha inicio - fecha finalizacion
								   	if l_fecpar <= CDate(l_calfecfin) and l_fecpar >= CDate(l_calfecini)  then 
										l_sql = "INSERT INTO cap_calendario "
				 			           	l_sql = l_sql & "(calfecha, caldia, calhordes, calhorhas, lugnro, evmonro ) "
   		                                l_sql = l_sql & "VALUES (" & cambiafecha(l_fecpar,"YMD",true)  & ",'" & l_caldia & "','" & l_calhordes1 & l_calhordes2 & "','" & l_calhorhas1 & l_calhorhas2 & "'," & l_lugnro & "," & l_evmonro & ")"
       		                            l_cm.CommandText = l_sql
                                   		cmExecute l_cm, l_sql, 0   
									end if
							end if	
					  end if	
					Next 
					if month(l_fini) = 12 then
			            l_fini = datevalue ("01/" & "01" & "/" + CStr(year(l_fini) + 1) )
			        else 
                        l_fini = datevalue ("01/" &  CStr(month(l_fini) + 1) & "/" + CStr(year(l_fini)))
					end if
				
				Wend
				'l_rs.close
End Select

l_sql = " SELECT evenro "
l_sql = l_sql & " FROM cap_eventomodulo "
   l_sql = l_sql & " WHERE evmonro = " & l_evmonro
rsOpen l_rs, cn, l_sql, 0
l_evento = l_rs("evenro")
if not l_rs.eof then ' Estoy posicionado en un Evento de Capacitacion
    l_rs.close
	Call FecIniFin (l_evento) ' Actualiza la Fecha de Inicio y Finalizacion del Evento
end if	

'asistente de evetos
 call actualizarPasos(32, l_evento, -1)
 
'l_rs.Close
Set l_rs = nothing
Set l_cm = Nothing
%>
 <script>
 opener.parent.ifrm.location.reload();
 opener.parent.ifrm2.location.reload();
 
 
 if (opener.parent.parent.parent.RefrescarPasos){
    opener.parent.parent.parent.RefrescarPasos();
 }
 </script>
<%
	Response.write "<script>alert('Operación Realizada.');window.close();</script>"
%>

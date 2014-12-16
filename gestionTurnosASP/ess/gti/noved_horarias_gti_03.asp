<% Option Explicit %>
<!--#include virtual="/serviciolocal/shared/inc/sec.inc"-->
<!--#include virtual="/serviciolocal/shared/inc/const.inc"-->
<!--#include virtual="/serviciolocal/shared/db/conn_db.inc"-->
<!--#include virtual="/serviciolocal/shared/inc/fecha.inc"-->
<!--#include virtual="/serviciolocal/shared/inc/sqls.inc"-->
<!--#include virtual="/serviciolocal/shared/inc/buscarturno.inc"-->
<!--#include virtual="/serviciolocal/shared/inc/buscardia.inc"-->

<!--
-----------------------------------------------------------------------------
Archivo        : noved_horarias_gti_03.asp
Descripcion    : Modulo que se encarga de guardar los datos de las nov horarias
Modificacion   :
    18/09/2003 - Scarpa D.   - Coordinacion con el tablero del empleado
    06/10/2003 - Scarpa D.   - Punto de procesamiento
    07/10/2003 - Scarpa D.   - Motivo no obligatorio	
    27/02/2004 - Muzzolón M. - Coustomizacion para Red Megatone		
	07/10/2005- Leticia A. - 
-----------------------------------------------------------------------------
-->
<html>
<head>
</head>
<body>
<script src="/serviciolocal/shared/js/fn_windows.js"></script>
<script src="/serviciolocal/shared/js/fn_ayuda.js"></script>
<script src="/serviciolocal/shared/js/fn_fechas.js"></script>
<% 
on error goto 0

const l_valornulo = "null"

Dim l_tipo
Dim l_cm
Dim l_sql
dim l_sql3
dim l_sql4

Dim l_gnovnro
Dim l_gnovdesabr
Dim l_gnovdesext
Dim l_gnovotoa
Dim l_gtnovnro
Dim l_motnro
Dim l_gnovtipo
Dim l_gnovdesde
Dim l_gnovhasta
Dim l_gnovhoradesde
Dim l_gnovhorahasta
Dim l_gnovorden
Dim l_gnovmaxhoras
Dim l_rs
Dim l_rs5
Dim myRs

Dim l_datos
Dim l_fechadesde
Dim l_fechahasta
Dim l_ternro
Dim username

Dim lngAlcanGrupo ' VER!!!!!!!!!!!
    lngAlcanGrupo = 2
	
const max_horas = 8 'cantidad de hs permitidad para un dia o una semana cuando el timpo de novedad es 1 (para red megatone)
	
Dim l_Cysfirmas
Dim l_Cysfirmas1

l_ternro = request.querystring("ternro")
l_fechadesde = request.querystring("fechadesde")
l_fechahasta  = request.querystring("fechahasta")

l_tipo = request.querystring("tipo")
l_gnovnro = request.Form("gnovnro")
l_gnovdesabr = request.Form("gnovdesabr") 
l_gnovdesext = request.Form("gnovdesext")
l_gnovotoa = request.Form("gnovotoa")
l_gtnovnro = request.Form("gtnovnro")
l_motnro = request.Form("motnro")
l_gnovdesde = request.Form("gnovdesde")
l_gnovhasta = request.Form("gnovhasta")
l_gnovtipo = request.Form("gnovtipo")

l_cysfirmas  = request.Form("seleccion")
l_cysfirmas1 = request.Form("seleccion1")

'----------------------------------------------------------------------------------------------------------
'FUNCION: imprime el dia de la semana
function imprDia(fecha)
Dim l_dia

select case Weekday(fecha)
		case 2
		l_dia = "Lunes"
		case 3
		l_dia = "Martes"
		case 4
		l_dia = "Miercoles"
		case 5
		l_dia = "Jueves"
		case 6
		l_dia = "Viernes"
		case 7
		l_dia = "Sabado"
		case 1
		l_dia = "Domingo"
end select
imprDia = l_dia
end function 'imprDia(fecha)

'----------------------------------------------------------------------------------------------------------
'FUNCION: busca el primer dia de la semana (lunes)
function primerDiaSem (dia)
if imprDia(dia) = "Lunes" then 
   primerDiaSem = dia
   else primerDiaSem = primerDiaSem(cstr(dateadd("d",-1,cdate(dia))))
end if
end function

'----------------------------------------------------------------------------------------------------------
'FUNCION: busca el ultimo dia de la semana (domingo)
function ultimoDiaSem (dia)
if imprDia(dia) = "Domingo" then 
   ultimoDiaSem = dia
   else ultimoDiaSem = ultimoDiaSem(cstr(dateadd("d",1,cdate(dia))))
end if
end function

'----------------------------------------------------------------------------------------------------------
'FUNCION: Convierte un string que contiene una hora al formato string hora
function convHora(stri)
  convHora = mid (stri,2,2) & ":" & mid (stri,4,2) 
end function 'convHora(str)

'----------------------------------------------------------------------------------------------------------

'FUNCION: Convierte un string que contiene una hora al formato string hora
function convHora2(stri)
  convHora2 = mid (stri,1,2) & ":" & mid (stri,3,2) 
end function 'convHora(str)


'----------------------------------------------------------------------------------------------------------

'FUNCION: obtiene el dia de una fecha
function dia(str)
  dia = mid (str,1,2) 
end function 'dia

'----------------------------------------------------------------------------------------------------------

'FUNCION: obtiene el mes de una fecha
function mes(str)
  mes = mid (str,4,2) 
end function 'dia

'----------------------------------------------------------------------------------------------------------
'FUNCION: obtiene el año de una fecha
function anio(str)
  anio = mid (str,7,4) 
end function 'dia

'----------------------------------------------------------------------------------------------------------


'FUNCION: Busca una novedad que se intersecte con las fechas desde hasta y retorna las horas acumuladas entre esas fechas.
function horasNovedad ( fechadesde, fechahasta)

dim l_rs3
dim l_rs4
dim l_sql
dim l_sql2

dim l_dia
dim l_salir
dim l_hs_acum_x_semana
dim l_horas_x_dia

Set l_rs3 = Server.CreateObject("ADODB.RecordSet")
Set l_rs4 = Server.CreateObject("ADODB.RecordSet")

horasNovedad  = 0

if fechadesde > fechahasta then 
  exit function 
end if

'Busco una novedad para ese empleado que se intersecte con las fechas.
l_sql = "SELECT gnovnro, gnovhasta, gnovdesde, gnovhoradesde , gnovhorahasta, gnovmaxhoras, gnovtipo  "
l_sql = l_sql & " FROM gti_novedad "
l_sql = l_sql & " WHERE gnovotoa = " & l_gnovotoa 
l_sql = l_sql & " and not (gnovhasta < " & cambiafecha(fechadesde,"YMD",true)
l_sql = l_sql & "          or gnovdesde > " & cambiafecha(fechahasta,"YMD",true) & ") "
l_sql = l_sql & " and gtnovnro = 1 " 
if (l_tipo ="M") then
	l_sql = l_sql & " and gnovnro <>" & l_gnovnro
end if

rsOpen l_rs3, cn, l_sql, 0 

do while not l_rs3.eof 
'Para cada novedad que se intersecta:
'Recorro todos los dias de la novedad y sumo las hs de los que estan entre las fechas de entrada

		l_dia   	= cstr(l_rs3("gnovdesde"))
		l_salir 	= false
		l_hs_acum_x_semana = 0
		l_horas_x_dia = 0
		
		select case l_rs3("gnovtipo")
		case 2  'Parcial Fijo
				l_horas_x_dia = DateDiff("n", convHora2(l_rs3("gnovhoradesde")), convHora2(l_rs3("gnovhorahasta")) ) / 60
		case 3 	'Parcial Variable
				l_horas_x_dia = cint(l_rs3("gnovmaxhoras"))
		end select

		'Recorro todos los dias de la novedad:
		do while not l_salir 
			if cdate(l_dia) >= cdate(fechadesde) and cdate(l_dia <= fechahasta) then 'considero el dia, esta entre las fechas.
				'Busco el turno del empleado
				BuscarTurno l_gnovotoa, l_dia
				if Tiene_Turno Then
				  ' Busco el sub-turno del empleado
 				    buscar_Dia l_dia,cdate(fecha_inicio),NroTurno,l_gnovotoa,P_asignacion
					if Nro_Dia <> -1 then 
						select case l_rs3("gnovtipo")
						case 1 'Dia Completo
						    'busco las horas minimas del dia 
						    l_sql2 = "SELECT gti_dias.diacanthoras FROM gti_dias WHERE gti_dias.dianro = " & Nro_dia
						    l_rs4.open l_sql2,cn
						    if not l_rs4.eof then 
			      			  l_hs_acum_x_semana = l_hs_acum_x_semana + csng(l_rs4("diacanthoras"))
						    end if
						    l_rs4.close
						    'Fin dia completo
						case 2,3 'Parcial Fijo o Variable
			 			 	if not blnDia_libre then 'el empleado no tiene dia libre, acumulo las horas  del dia
								l_hs_acum_x_semana = l_hs_acum_x_semana + l_horas_x_dia
							end if
						 end select
					end if'nro dia <> -1
				end if'tiene turno
			end if  'l_dia entre fechas 
			
			if l_dia = cstr(l_rs3("gnovhasta"))  then 'termino la novedad 
				l_salir = true
			else ' Paso al proximo dia
				l_dia = cstr(dateadd("d",1,cdate(l_dia)))
			end if	
		loop 'Fin de recorrer todos los dias de una novedad
	horasNovedad = l_hs_acum_x_semana + horasNovedad

	l_rs3.MoveNext	
loop 'cada novedad que se intersecta

end function 

'----------------------------------------------------------------------------------------------------------
'FUNCION : hace las validaciones requeridas para red megatone.
function  valida_acum_hs

'Coustumizacion Red Megatone: para tipos de novedades (1)-Compenzacion Red Megatone.
'Caso 1: la novedad no debe superar en  ninguno de sus dias las 8 hs.
'Caso 2: la suma de las hs de una novedad en una semana no debe superar las 8 hs.
'Caso 3: la suma de las hs de una novedad en una semana mas otra novedad en la misma semana no debe superar las 8 hs.

dim l_rs2

dim l_dia
dim l_salir
dim l_sql2
dim l_excede_hs
dim l_hs_acum_x_semana
dim l_pri_dia_semana
dim l_ult_dia_semana
dim l_caso1
dim l_caso2
dim l_caso3
dim l_hs_X_dia

l_dia   	= l_gnovdesde
l_salir 	= false
l_excede_hs = false
l_caso1		= false
l_caso2		= false
l_caso3		= false
l_hs_acum_x_semana = 0
l_pri_dia_semana = primerDiaSem(l_gnovdesde)
l_ult_dia_semana = ultimoDiaSem(l_gnovhasta)

if l_gnovtipo = 2 then 'PARCIAL FIJO
	l_hs_X_dia = DateDiff("n", convHora(l_gnovhoradesde), convHora(l_gnovhorahasta)) / 60
	else 
		if l_gnovtipo = 3 then 'PARCIAL VARIABLE
			l_hs_X_dia = cint(l_gnovmaxhoras)
		end if
end if

'Controlo dia por dia de la novedad que el empleado no tenga mas de max_horas hs.
do while not l_salir and not l_excede_hs
	
	'Busco el turno del empleado
	BuscarTurno l_gnovotoa, l_dia
	if Tiene_Turno Then
	  ' Busco el sub-turno del empleado y si tiene dia libre
		    buscar_Dia l_dia,cdate(fecha_inicio),NroTurno,l_gnovotoa,P_asignacion
			
		if Nro_Dia <> -1 then 
		   select case l_gnovtipo 
		   
		   case 1 'DIA COMPLETO
		   	  Set l_rs2 = Server.CreateObject("ADODB.RecordSet")
			  'busco las horas minimas del dia 
			   l_sql2 = "SELECT gti_dias.diacanthoras FROM gti_dias WHERE gti_dias.dianro = " & Nro_dia
			   l_rs2.open l_sql2,cn
			   if not l_rs2.eof then 
				  if csng(l_rs2("diacanthoras")) > max_horas then 'caso 1
				    l_caso1		= true
					l_excede_hs = true
				  end if
				  
				 'acumulo las hs semanales:
				 if imprDia(l_dia) = "Lunes" then
				 	if horasNovedad(l_pri_dia_semana, cstr(dateadd("d",-1,cdate(l_dia))) ) + l_hs_acum_x_semana > max_horas then 'CASO 3
						l_caso3		= true
						l_excede_hs = true
					end if
					l_pri_dia_semana   = l_dia
				 	l_hs_acum_x_semana = csng(l_rs2("diacanthoras"))
				 else
				 	l_hs_acum_x_semana = l_hs_acum_x_semana + csng(l_rs2("diacanthoras"))
				 end if 
				
				 if l_hs_acum_x_semana > max_horas then 'caso2
					l_caso2 	= true
			  		l_excede_hs = true
				 end if  
		   	   l_rs2.close 
			   end if ' existe el dia
  				'FIN DIA COMPLETO.
		   	
		   case 2 , 3	' PARCIAL FIJO: si el empleado no tiene dia libre, se considera hs de trabajo: hora desde - hara hasta.
				   		' PARCIAL VARIABLE: si el empleado no tiene dia libre, se considera hs de trabajo: maximo de hs.
		   		
				if not blnDia_libre then 'el empleado no tiene dia libre
	  				if l_hs_X_dia > max_horas then 'caso 1
	    				l_caso1		= true
						l_excede_hs = true
	  				end if
	  				
	 				'acumulo las hs semanales:
	 				if imprDia(l_dia) = "Lunes" then
	 					if horasNovedad(l_pri_dia_semana, cstr(dateadd("d",-1,cdate(l_dia))) ) + l_hs_acum_x_semana > max_horas then 'CASO 3
							l_caso3		= true
							l_excede_hs = true
						end if
						l_pri_dia_semana   = l_dia
	 					l_hs_acum_x_semana = l_hs_X_dia
	 				else
	 					l_hs_acum_x_semana = l_hs_acum_x_semana + l_hs_X_dia
	 				end if 
					
	 				if l_hs_acum_x_semana > max_horas then 'caso2
						l_caso2 	= true
  						l_excede_hs = true
	 				end if  
							
				end if 'dia libre
			   'FIN PARCIAL FIJO O VARIABLE
			
			end select
			
		end if'nro dia <> -1
	end if'tiene turno
	
	if l_dia = l_gnovhasta then 'termino la novedad
		l_salir = true
	else  ' Paso al proximo dia
		l_dia = cstr(dateadd("d",1,cdate(l_dia)))
	end if	
	
loop 'NO EXEDE HS NI CONSIDERO TODOS LOS DIAS DE LA NOVEDAD

if l_salir then 'No exede hs.
   	if horasNovedad(l_pri_dia_semana, l_ult_dia_semana) + l_hs_acum_x_semana > max_horas then 'CASO 3
		l_caso3		 = true
		l_excede_hs  = true
	end if
end if

if l_caso1 then 
 response.write "<script>alert(' Un dia tiene más de " &  max_horas  & " hs.   ');history.back();</script> "
end if
if l_caso2 then 
 response.write "<script>alert(' Acumula más de " & max_horas & " hs en la semana.  ');history.back();</script>"
end if
if l_caso3 then 
 response.write "<script>alert(' Hay otra novedad que junto con ésta suman mas de " & max_horas & " hs. en una semana. ');history.back();</script>" 
end if
valida_acum_hs = not (l_caso1 or l_caso2 or l_caso3)

end function

'----------------------------------------------------------------------------------------------------------
'COMIENZO

if l_motnro = "" then
   l_motnro = "null"
end if

if l_gnovtipo = 1 then
	l_gnovhoradesde = l_valornulo
	l_gnovhorahasta = l_valornulo
	l_gnovorden = l_valornulo
	l_gnovmaxhoras = l_valornulo
end if
if l_gnovtipo = 2 then 
	l_gnovhoradesde = "'"&request.Form("gnovhoradesde1") & request.Form("gnovhoradesde2")&"'"
	l_gnovhorahasta = "'"&request.Form("gnovhorahasta1") & request.Form("gnovhorahasta2")&"'"
	l_gnovorden = l_valornulo
	l_gnovmaxhoras = l_valornulo
end if 
if l_gnovtipo = 3 then 
	l_gnovhoradesde = l_valornulo
	l_gnovhorahasta = l_valornulo
	l_gnovorden = request.Form("gnovorden")
	l_gnovmaxhoras = request.Form("gnovmaxhoras")
end if 

'controlamos que no haya otro registro con esa configuracion
Set l_rs = Server.CreateObject("ADODB.RecordSet")
l_sql = "SELECT gnovnro "
l_sql = l_sql & " FROM gti_novedad "
l_sql = l_sql & " WHERE gnovotoa="& l_gnovotoa &" and ((gnovdesde >=" & cambiafecha(l_gnovdesde,"YMD",true)
l_sql = l_sql & " and gnovdesde <=" & cambiafecha(l_gnovhasta,"YMD",true) & ") "
l_sql = l_sql & " or (gnovhasta >=" & cambiafecha(l_gnovdesde,"YMD",true)
l_sql = l_sql & " and gnovhasta <=" & cambiafecha(l_gnovhasta,"YMD",true) & ") "
l_sql = l_sql & " or (gnovdesde <=" & cambiafecha(l_gnovdesde,"YMD",true)
l_sql = l_sql & " and gnovhasta >=" & cambiafecha(l_gnovdesde,"YMD",true) & ") "
l_sql = l_sql & " or (gnovdesde <=" & cambiafecha(l_gnovhasta,"YMD",true)
l_sql = l_sql & " and gnovhasta >=" & cambiafecha(l_gnovhasta,"YMD",true) & ")) "
if (l_tipo ="M") then
	l_sql = l_sql & " and gnovnro <>" & l_gnovnro
end if
if l_gnovtipo=2 then
	l_sql = l_sql & " and (gnovtipo =1 " 
	l_sql = l_sql & " or ( gnovtipo=2 and ( (gnovhoradesde <=" & l_gnovhorahasta & " and gnovhorahasta>="& l_gnovhorahasta & " ) "
	l_sql = l_sql & " or (gnovhoradesde <=" & l_gnovhoradesde & " and gnovhorahasta>="& l_gnovhoradesde & " ) "
	l_sql = l_sql & " or (gnovhoradesde >=" & l_gnovhoradesde & " and gnovhoradesde<="& l_gnovhorahasta & " )))) "
end if
if l_gnovtipo=3 then
	l_sql = l_sql & " and gnovtipo =1 " 
end if
rsOpen l_rs, cn, l_sql, 0 

if not l_rs.eof then
	Response.write "<script>alert('Esta Licencia se superpone con otras cargadas anteriormente.');history.back();</script>"
	l_rs.close
else	


	'-------------Coustumizacion Red Megatone--------------------------------------------------------------'
	'--                                                                                                 ---'
		if l_gtnovnro = 1 then 
			
			'Obtengo el puesto del empleado al que se le esta queriendo generar la novedad
			Set l_rs5 = Server.CreateObject("ADODB.RecordSet")
			l_sql4 = " SELECT estrnro FROM his_estructura " 
			l_sql4 = l_sql4 & " WHERE  tenro = 4 AND ternro = " & l_gnovotoa  
			l_sql4 = l_sql4 & " AND htethasta IS NULL "
			
			rsOpen l_rs5, cn, l_sql4, 0
			
			if not l_rs5.eof then 'tiene puesto
			    if dia(l_gnovdesde) >= "01" and dia(l_gnovhasta) <= "15" and mes(l_gnovdesde) = mes(l_gnovhasta)   and anio(l_gnovdesde) = anio(l_gnovhasta) and  INstr ( "632,659,1640" , l_rs5("estrnro")) > 0   then 
		  			'obtengo el el perfil del usuario logueado
					Set myRs = Server.CreateObject("ADODB.RecordSet")
					username = UCase(Session("Username"))
					l_sql3 = "SELECT perfnom FROM user_per "
					l_sql3 = l_sql3 & " inner join perf_usr on perf_usr.perfnro = user_per.perfnro"
					l_sql3 = l_sql3 & " where upper(iduser) = '" & username & "'"
					rsOpen myRs, cn, l_sql3, 0
					if myRs("perfnom") <> "Administrador" then 'en las fechas entre 1 y 15 si el puesto es gerente o cajero solo pueden cargar una novedad un Administrador
						response.write "<script>alert('El Usuario no puede Cargar la Novedad Horaria.');window.close();</script>"
						response.end
					else 'es administrador, las fechas estan entre 1 y 15 y es gerente o cajero, si valida puede cargar la novedad
						if not valida_acum_hs then 'Aborto
						   response.end
						end if
					end if
					myRs.Close
				else ' o no es ni gerente o cajero o las fechas no estan entre el 1 y el 15, si valida puede cargar la novedad
					if not valida_acum_hs then 'Aborto
					    response.end
					end if
			   	end if
			else 'no tiene puesto, Aborto
		   		  response.write "<script>alert('El Empleado no tiene un Puesto.');window.close();</script>"
				  response.end
			end if 'tiene puesto
			l_rs5.close
		end if 'tipo de novedad = 1
	'--                                                                                                               ---'
	'-------------FIN Coustumizacion Red Megatone------------------------------------------------------------------------'


		l_rs.close
		cn.beginTrans
		set l_cm = Server.CreateObject("ADODB.Command")
		if l_tipo = "A" OR l_tipo = "TA" then 
			l_sql = "insert into gti_novedad "
			l_sql = l_sql & "( gnovdesabr, gnovdesext, gnovotoa, gtnovnro, motnro, gnovtipo "
			l_sql = l_sql & ", gnovdesde, gnovhasta, gnovhoradesde, gnovhorahasta, gnovorden, gnovmaxhoras, gnovestado)"
			l_sql = l_sql & "values ('" & l_gnovdesabr &"', '" & l_gnovdesext & " ', " & l_gnovotoa & ", " & l_gtnovnro & ", " & l_motnro & ", " & l_gnovtipo
			l_sql = l_sql & ", " & cambiafecha(l_gnovdesde,"YMD",true)  & ", " & cambiafecha (l_gnovhasta,"YMD",true) & ", " & l_gnovhoradesde & ", " & l_gnovhorahasta 
			l_sql = l_sql & ", " & l_gnovorden & ", " & l_gnovmaxhoras & ",' ')"
		else
			l_sql = "update gti_novedad "
			l_sql = l_sql & "set  gnovdesabr='"& l_gnovdesabr & "', gnovdesext='" & l_gnovdesext & "  ', gnovotoa =" & l_gnovotoa & ", motnro =" & l_motnro & ", gtnovnro =" & l_gtnovnro 
			l_sql = l_sql & ", gnovtipo =" & l_gnovtipo & ", gnovdesde="& cambiafecha(l_gnovdesde,"YMD",true) & ", gnovhasta =" & cambiafecha(l_gnovhasta,"YMD",true) 
			l_sql = l_sql & ", gnovhoradesde = " & l_gnovhoradesde & ", gnovhorahasta=" & l_gnovhorahasta & ", gnovmaxhoras = " & l_gnovmaxhoras & ", gnovorden =" & l_gnovorden 
			l_sql = l_sql & " where gnovnro = " & l_gnovnro
		end if
		l_cm.activeconnection = Cn
		l_cm.CommandText = l_sql
		cmExecute l_cm, l_sql, 0

		'ingresamos la justificacion
		if l_tipo = "A" OR l_tipo = "TA" then 
			Set l_rs = Server.CreateObject("ADODB.RecordSet")
			'l_sql = "select unique DBINFO('sqlca.sqlerrd1') as nov_id from gti_novedad "
			l_sql = fsql_seqvalue("nov_id","gti_novedad")
			l_rs.MaxRecords = 1
			rsOpen l_rs, cn, l_sql, 0
			l_gnovnro=l_rs("nov_id")
			l_rs.Close
			
		    l_sql = "INSERT INTO gti_justificacion ( jusanterior,juscodext,jusdesde,jusdiacompleto,jushasta,jussigla,jussistema,ternro,tjusnro,turnro,jushoradesde,jushorahasta,juseltipo,juselorden,juselmaxhoras )" & _
		                    " VALUES( -1," & l_gnovnro & "," & cambiafecha(l_gnovdesde,"YMD",true) & ",-1," & cambiafecha(l_gnovhasta,"YMD",true) & ",'NOV',-1," & l_gnovotoa & ",1,0," & l_gnovhoradesde & "," & l_gnovhorahasta & "," & l_gnovtipo & "," & l_gnovorden & "," & l_gnovmaxhoras & ")"
			l_cm.activeconnection = Cn
			l_cm.CommandText = l_sql
			cmExecute l_cm, l_sql, 0
		
		else
			l_sql = "UPDATE gti_justificacion SET jusanterior = -1, jusdesde = " & cambiafecha(l_gnovdesde,"YMD",true) & ", jusdiacompleto = -1, jushasta = " & cambiafecha(l_gnovhasta,"YMD",true) & ",jussistema = -1"  & _
		                      ", tjusnro = 1 ,turnro = 0 ,jushoradesde = " & l_gnovhoradesde & ", jushorahasta = " & l_gnovhorahasta & ", juseltipo = " & l_gnovtipo & ", juselorden = " & l_gnovorden & ", juselmaxhoras = " & l_gnovmaxhoras	&_
							 " WHERE ternro= " & l_gnovotoa & " and jussigla = 'NOV' and juscodext = " & l_gnovnro
							 
			l_cm.activeconnection = Cn
			l_cm.CommandText = l_sql
			cmExecute l_cm, l_sql, 0
		end if
		
		'circuito de firmas
		if (l_cysfirmas1 <> "") then
		  if inStr(l_cysfirmas1,"@@@") <> 0 then
		    l_cysfirmas1 = left(l_cysfirmas1,inStr(l_cysfirmas1,"@@@")-1) & l_gnovnro & mid(l_cysfirmas1,inStr(l_cysfirmas1,"@@@")+3)
		  end if
		  l_cm.activeconnection = Cn
		  l_cm.CommandText = l_cysfirmas1
		  cmExecute l_cm, l_cysfirmas1, 0
		end if  
		
		if (l_cysfirmas <> "") then
		  if inStr(l_cysfirmas,"@@@") <> 0 then
		    l_cysfirmas = left(l_cysfirmas,inStr(l_cysfirmas,"@@@")-1) & l_gnovnro & mid(l_cysfirmas,inStr(l_cysfirmas,"@@@")+3)
		  end if
		  l_cm.activeconnection = Cn
		  l_cm.CommandText = l_cysfirmas
		  cmExecute l_cm, l_cysfirmas, 0
		end if  
		
		cn.CommitTrans
		
		Set cn = Nothing
		Set l_cm = Nothing
		l_datos = "ternro="& l_gnovotoa & "  "
		
		
		if ((esmenor(l_gnovdesde,l_fechahasta) or  (l_gnovdesde=l_fechahasta)) and (esmenor(l_fechadesde,l_gnovdesde) or  (l_gnovdesde=l_fechadesde))) then
			l_datos =l_datos & "&fechadesde="& l_fechadesde & "&fechahasta="& l_fechahasta
		else
			l_datos =l_datos & "&fechadesde="& l_gnovdesde & "&fechahasta="& l_gnovdesde
		end if 
		
	   Response.write "<script>window.opener.ifrm.location.reload();</script>"
%>

<script>
  alert('Operación Realizada.');
  window.close();
</script>

<% 
end if
%>

</body>
</html>


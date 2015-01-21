<%@ Language=VBScript %>
<% Option Explicit %>
<!--#include virtual="/turnos/shared/inc/sec.inc"-->
<!--#include virtual="/turnos/shared/inc/const.inc"-->
<!--#include virtual="/turnos/shared/db/conn_db.inc"-->
<!--#include virtual="/turnos/shared/inc/fecha.inc"-->



<head>

<meta http-equiv="Content-Type" content="text/html; charset=iso-8859-1">
<script src="/turnos/shared/js/fn_windows.js"></script>
<script src="/turnos/shared/js/fn_confirm.js"></script>
<script src="/turnos/shared/js/fn_ayuda.js"></script>
<script src="/turnos/shared/js/fn_fechas.js"></script>
</head>



<%
'pagina que recibirá la fecha seleccionada por el usuario
Const URLDestino = "AsignarTurnos_con_00.asp" 

Dim MyMonth 'Month of calendar
Dim MyYear 'Year of calendar
Dim FirstDay 'First day of the month. 1 = Monday
Dim CurrentDay 'Used to print dates in calendar
Dim Col 'Calendar column
Dim Row 'Calendar row
Dim l_referencia

Dim l_id

Dim l_hd
Dim l_md
Dim l_hh
Dim l_mh
Dim l_horadesde
Dim l_horahasta

MyMonth = Request.Querystring("Month")
MyYear = Request.Querystring("Year")
l_id = Request.Querystring("id")

l_hd = Request.Querystring("hd")
l_md = Request.Querystring("md")
l_hh = Request.Querystring("hh")
l_mh = Request.Querystring("mh")




l_horadesde = l_hd & ":" & l_md
l_horahasta = l_hh & ":" & l_mh

'response.write l_horadesde
'response.write l_horahasta


If IsEmpty(MyMonth) then MyMonth = Month(Date)
if IsEmpty(MyYear) then MyYear = Year(Date)



Call ShowHeader (MyMonth, MyYear)

FirstDay = WeekDay(DateSerial(MyYear, MyMonth, 1)) -1
CurrentDay = 1

'Let's build the calendar
For Row = 0 to 5
response.write "<tr>"
	For Col = 0 to 6
		If Row = 0 and Col < FirstDay then
			response.write "<td> </td>"
		elseif CurrentDay > LastDay(MyMonth, MyYear) then
			response.write "<td> </td>"
		else
			response.write "<td "  ' onclick=Javascript:parent.abrirVentana('templatereservas_con_02.asp,'',520,200);
			if cInt(MyYear) = Year(Date) and cInt(MyMonth) = Month(Date) and CurrentDay = Day(Date) then 
				response.write " class='calCeldaResaltado' align='center'>"
			else 
				response.write " align='center'>"
			end if
			
			l_referencia = Referencia (CurrentDay, MyMonth, MyYear)
			
			if l_referencia = "blanco" then
				response.write "<a>" 			
			else
				response.write "<a   target='_blank' href='" & URLDestino & "?id=" & l_id & "&day=" & CurrentDay _
							& "&month=" & MyMonth & "&year=" & MyYear & "'>" 
			end if
			
			Response.Write "<div class='" & l_referencia & "'>" 
				
						
			'if cInt(MyYear) = Year(Date) and cInt(MyMonth) = Month(Date) and CurrentDay = Day(Date) then 
			'	Response.Write "<div class='calResaltado'>" 
			'else
			'	Response.Write "<div class='calSimbolo'>" 
			'end if
			
			Response.Write CurrentDay  & "</div></a></td>"
			'Response.Write  Referencia (CurrentDay, MyMonth, MyYear)
			CurrentDay = CurrentDay + 1
		End If
	Next
	response.write "</tr>"
Next
response.write "</table>"

Call SignificadoReferencia

response.write "</body></html>"


'------ Sub and functions


Sub ShowHeader(MyMonth,MyYear)
%>
<html>
<head>
	<meta http-equiv='Content-Type' content='text/html; charset=iso-8859-1'>
	<style>
		.calFondoCalendario {background-color:white}
		.calEncabe {font-family:Arial, Helvetica, sans-serif; font-size:15px}
		.calFondoEncabe {background-color:D0D2D5}
		.calDias {font-family:Arial, Helvetica, sans-serif; font-size:12px; font-weight:900}
		.calSimbolo {font-family:Arial, Helvetica, sans-serif; font-size:11px; text-decoration:none; font-weight:200; color:blue}
		.calResaltado {font-family:Arial, Helvetica, sans-serif; font-size:11px; text-decoration:none; font-weight:700}
		.calCeldaResaltado {background-color:lightyellow}
		.Verde {font-family:Arial, Helvetica, sans-serif; font-size:13px; text-decoration:none; font-weight:200; color:blue ;  background-color:lightgreen}
		.Rojo {font-family:Arial, Helvetica, sans-serif; font-size:13px; text-decoration:none; font-weight:700; background-color:red}
		.Blanco {font-family:Arial, Helvetica, sans-serif; font-size:13px; text-decoration:none; font-weight:700; background-color:white}
  	</style>
</head>

<body bgcolor='CDCDCD'>


<table  border='1' cellspacing='' cellpadding='3' width='250' align='center' class="calFondoCalendario">
	<tr align='center'> 
		<td colspan='7'>
			<table border='0' cellspacing='1' cellpadding='1' width='100%' class="calFondoEncabe">
				<tr>
					<td align='left'>
						<%
						response.write "<a href = 'Agenda_con_01.asp?id=" & l_id & "&"
						if MyMonth - 1 = 0 then 
							response.write "month=12&year=" & MyYear -1
						else 
							response.write "month=" & MyMonth - 1 & "&year=" & MyYear
						end if
						response.write "'><span class='calSimbolo'><<</span></a>"

						response.write "<span class='calEncabe'> <b>" & MonthName(MyMonth) & "</b> </span>"

						response.write "<a href = 'Agenda_con_01.asp?id=" & l_id & "&"
						if MyMonth + 1 = 13 then 
							response.write "month=1&year=" & MyYear + 1
						else 
							response.write "month=" & MyMonth + 1 & "&year=" & MyYear
						end if
						response.write "'><span class='calSimbolo'>>></span></a>"
						%>
					</td>
					<td align='center'>
						<%
						response.write "<a href = 'Agenda_con_01.asp?id=" & l_id & "&"
						response.write "month=" & Month(Date()) & "&year=" & Year(Date())
						response.write "'><div class='calSimbolo'><b>Hoy</b></div></a>"
						%>						
					</td>
					<td align='right'>
						<%
						response.write "<a href = 'Agenda_con_01.asp?id=" & l_id & "&"
						response.write "month=" & MyMonth & "&year=" & MyYear -1
						response.write "'><span class='calSimbolo'><<</span></a>"

						response.write "<span class='calEncabe'> <b>" & MyYear & "</b> </span>"
						response.write "<a href = 'Agenda_con_01.asp?id=" & l_id & "&"
						response.write "month=" & MyMonth & "&year=" & MyYear + 1
						response.write "'><span class='calSimbolo'>>></span></a>"
						%>
					</td>
				</tr>
			</table>
		</td>
	</tr>
	<tr align='center'> 
		<td><div class='calDias'>D</div></td>
		<td><div class='calDias'>L</div></td>
		<td><div class='calDias'>M</div></td>
		<td><div class='calDias'>X</div></td>
		<td><div class='calDias'>J</div></td>
		<td><div class='calDias'>V</div></td>
		<td><div class='calDias'>S</div></td>
	</tr>
<%
End Sub

Sub SignificadoReferencia()
%>

<table  border='1' cellspacing='' cellpadding='3' width='250' align='center' class="calFondoCalendario">
	<tr align='center'> 
		<td colspan='7'>
			<table border='0' cellspacing='1' cellpadding='1' width='100%' class="calFondonEncabe">
				<tr>
					<td align='left'>
				    <div class='Verde'>Verde</div>
					</td>
					<td align='center'>
					Turnos Disponibles				
					</td>
				</tr>
				<tr>
					<td align='left'>
				    <div class='Rojo'>Rojo</div>
					</td>
					<td align='center'>
					Ningun Turno				
					</td>
				</tr>				
				<tr>
					<td align='left'>
				    <div class='Blanco'>Blanco</div>
					</td>
					<td align='center'>
					No Atiende				
					</td>
				</tr>				
			</table>
		</td>
	</tr>
	
<%
End Sub



Function MonthName(MyMonth)
	Select Case MyMonth
		Case 1
			MonthName = "Enero"
		Case 2
			MonthName = "Febrero"
		Case 3
			MonthName = "Marzo"
		Case 4
			MonthName = "Abril"
		Case 5
			MonthName = "Mayo"
		Case 6
			MonthName = "Junio"
		Case 7
			MonthName = "Julio"
		Case 8
			MonthName = "Agosto"
		Case 9
			MonthName = "Septiembre"
		Case 10
			MonthName = "Octubre"
		Case 11
			MonthName = "Noviembre"
		Case 12
			MonthName = "Diciembre"
		Case Else
			MonthName = "ERROR!"
	End Select
End Function

Function LastDay(MyMonth, MyYear)
' Returns the last day of the month. Takes into account leap years
' Usage: LastDay(Month, Year)
' Example: LastDay(12,2000) or LastDay(12) or Lastday


	Select Case MyMonth
		Case 1, 3, 5, 7, 8, 10, 12
			LastDay = 31

		Case 4, 6, 9, 11
			LastDay = 30

		Case 2
			If IsDate(MyYear & "-" & MyMonth & "-" & "29") Then LastDay = 29 Else LastDay = 28

		Case Else
			LastDay = 0
	End Select
End Function


Function Referencia (p_date, p_Month, p_Year)

Dim l_rs
Dim l_sql
Dim Fec
Dim CantCalendarios
Dim CantCalendariosLibres
Dim p_Monthcom

'response.write p_Month

if cint(p_date) < 10 then
 	p_date = "0" & p_date
end if  
if cint(p_Month) < 10 then
 	p_Monthcom = "0" & p_Month
else
	p_Monthcom = p_Month	
end if  
Fec = cstr(p_Monthcom) & "/" & cstr(p_date) & "/" & cstr(p_Year)
Fec = cstr(p_date) & "/" & cstr(p_Monthcom) & "/" & cstr(p_Year)
'response.write fec

'response.write p_date & " " & p_Month & " " & p_Year
'response.write cambiafecha( "10/25/2014",true,1)

Set l_rs = Server.CreateObject("ADODB.RecordSet")
l_sql = "SELECT  count(*) "
l_sql  = l_sql  & " FROM calendarios  "
'l_sql  = l_sql  & " WHERE CONVERT(VARCHAR(10), calendarios.fechahorainicio, 101)  = '" & Fec & "'"

l_sql = l_sql & " WHERE fechahorainicio >=" & cambiaformato (Fec,l_horadesde )
l_sql = l_sql & " AND fechahorainicio<=" & cambiaformato (Fec,l_horahasta )

l_sql  = l_sql  & " and calendarios.estado = 'ACTIVO' AND calendarios.idrecursoreservable = " & l_id


'response.write l_sql
rsOpen l_rs, cn, l_sql, 0
CantCalendarios = l_rs(0)
'response.write "can" &  CantCalendarios
l_rs.Close 


l_sql = "SELECT  count(*) "
l_sql  = l_sql  & " FROM calendarios  "
'l_sql  = l_sql  & " WHERE CONVERT(VARCHAR(10), calendarios.fechahorainicio, 101)  = '" & Fec & "'"
l_sql = l_sql & " WHERE fechahorainicio >=" & cambiaformato (Fec,l_horadesde )
l_sql = l_sql & " AND fechahorainicio<=" & cambiaformato (Fec,l_horahasta )
l_sql  = l_sql  & " AND calendarios.id not in (SELECT idcalendario FROM turnos ) "
l_sql  = l_sql  & " AND calendarios.estado = 'ACTIVO' AND calendarios.idrecursoreservable = " & l_id
rsOpen l_rs, cn, l_sql, 0
CantCalendariosLibres = l_rs(0)
'response.write "lib" &  CantCalendariosLibres
l_rs.Close 

 'response.write "ra" &  CantCalendarios & " " & CantCalendariosLibres

if clng(CantCalendarios) = 0 then
	Referencia = "blanco"
else
	if clng(CantCalendarios) > 0 and  clng(CantCalendariosLibres) = 0  then
		Referencia = "Rojo"
	else
		Referencia = "Verde"
	end if 	
end if 



'if clng(CantCalendarios) > 0 and clng(CantCalendariosLibres) >= 0 then
'	Referencia = "Verde"
'end if 



End Function








%>
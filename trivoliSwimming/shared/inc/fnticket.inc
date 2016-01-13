<%
'
'-----------------------------------------------------------------------------
'Archivo        : fnbuques.inc
'Descripcion    : Funciones varias a usar en los reportes de buques
'Creador        : Raul Chinestra
'Fecha Creacion : 07/04/2005
'-----------------------------------------------------------------------------

function kilosMermas(nromov)
' Datos Entrada: Un numero de movimiento
' Datos Salida: Calcula la sumatoria de los kilos de mermas que se le efectuaron en el calado

Dim l_rs
Dim l_sql
Dim Aux

Set l_rs = Server.CreateObject("ADODB.RecordSet")
	
	l_sql = " select rescalkilmer "
	l_sql = l_sql & " from tkt_calado "
	l_sql = l_sql & " inner join tkt_rescal on tkt_rescal.calnro = tkt_calado.calnro "
	l_sql = l_sql & " where tkt_calado.movnro = " & nromov
	rsOpen l_rs, cn, l_sql, 0
	Aux = 0
	do while not l_rs.eof 
		if not (isnull(l_rs(0))) then
			Aux = Aux + clng(l_rs(0))
		end if	
		l_rs.movenext
	loop
	kilosMermas = Aux

		
l_rs.close

end function 'KilosMermas


Function Bruto(nromov)
' Datos Entrada: Un numero de Movimiento
' Datos Salida: El Peso Bruto asociado al Movimiento, 0 en caso de no tenerlo

Dim l_rs
Dim l_sql

Set l_rs = Server.CreateObject("ADODB.RecordSet")
	l_sql = " SELECT peskg "
	l_sql = l_sql & " FROM tkt_pesaje "
	l_sql = l_sql & " WHERE tkt_pesaje.movnro = " & nromov & " AND tkt_pesaje.pestip = 'B'" 
	rsOpen l_rs, cn, l_sql, 0
	If Not l_rs.eof Then
		Bruto = l_rs("peskg")
	else 	
		Bruto = 0
	End If
l_rs.close

end function


Function Tara(nromov)
' Datos Entrada: Un numero de Movimiento
' Datos Salida: El Peso de Tara asociado al Movimiento, 0 en caso de no tenerlo

Dim l_rs
Dim l_sql

Set l_rs = Server.CreateObject("ADODB.RecordSet")
	l_sql = " SELECT peskg "
	l_sql = l_sql & " FROM tkt_pesaje "
	l_sql = l_sql & " WHERE tkt_pesaje.movnro = " & nromov & " AND tkt_pesaje.pestip = 'T'" 
	rsOpen l_rs, cn, l_sql, 0
	If Not l_rs.eof Then
		Tara = l_rs("peskg")
	else 	
		Tara = 0
	End If
l_rs.close

end function

Function NetoSMermas(nromov)
' Datos Entrada: Un numero de Movimiento
' Datos Salida: El Peso Neto sin Mermas calculado de la siguiente forma
' Bruto - Tara = NetoSMermas

	NetoSMermas = Clng(Bruto(nromov)) - Clng(Tara(nromov))

end function


function Neto(nromov)
' Datos Entrada: Un numero de movimiento
' Datos Salida: Calcula el Neto haciendo el siguiente calculo
' Peso Bruto - Peso Tara = Neto s/merma
' Neto s/merma - KilosMermas = Neto

	if isnull(kilosMermas(nromov)) then  
		Neto = clng(NetoSMermas(nromov))
	else
		Neto = clng(NetoSMermas(nromov)) - clng(kilosMermas(nromov))
	end if	

end function 'KilosMermas


function ValorObservado(nromov,rubro)
' Datos Entrada: Un numero de movimiento, codigo interno del Rubro
' Datos Salida: Calcula el Valor Observado para ese Rubro en el Movimiento

Dim l_rs
Dim l_sql

Set l_rs = Server.CreateObject("ADODB.RecordSet")
	l_sql = " select rescalvalobs "
	l_sql = l_sql & " from tkt_calado "
	l_sql = l_sql & " inner join tkt_rescal on tkt_rescal.calnro = tkt_calado.calnro and tkt_rescal.rubnro = " & rubro
	l_sql = l_sql & " where tkt_calado.movnro = " & nromov
	rsOpen l_rs, cn, l_sql, 0
	If Not l_rs.eof Then
		ValorObservado = l_rs(0)
	else 	
		ValorObservado = 0
	end if
	l_rs.close
	
end function 'ValorObservado


function Condicionales(nromov)
' Datos Entrada: Un numero de movimiento
' Datos Salida: Una descripción con todos los condicionales que aparecen en la etiqueta

Dim l_rs
Dim l_sql
Dim Aux

Set l_rs = Server.CreateObject("ADODB.RecordSet")
	l_sql = " select rubabr "
	l_sql = l_sql & " from tkt_calado "
	l_sql = l_sql & " inner join tkt_rescal on tkt_rescal.calnro = tkt_calado.calnro and tkt_rescal.rescalcon = -1 "
	l_sql = l_sql & " inner join tkt_rubro on tkt_rubro.rubnro = tkt_rescal.rubnro "	
	l_sql = l_sql & " where tkt_calado.movnro = " & nromov
	rsOpen l_rs, cn, l_sql, 0
	Aux = ""
	do while Not l_rs.eof
		Aux = Aux & " - " & l_rs(0)
		l_rs.movenext
	loop
	l_rs.close
	Condicionales = Aux
	
end function 'ValorObservado




Function MotivoForzado(nromov)
' Datos Entrada: Un numero de Movimiento
' Datos Salida: El Codigo del motivo de Rechazo
' Busca los codigo de los Motivos de Forzados si los hubiera de los pesos de tara y bruto y los separa por un guion

Dim l_rs
Dim l_sql
Dim AuxB
Dim AuxT

Set l_rs = Server.CreateObject("ADODB.RecordSet")
	l_sql = " SELECT pestip,motfornro "
	l_sql = l_sql & " FROM tkt_pesaje "
	l_sql = l_sql & " WHERE tkt_pesaje.movnro = " & nromov '& " AND 
	l_sql = l_sql & " "
	rsOpen l_rs, cn, l_sql, 0
	AuxB = ""
	AuxT = ""	
	do while not l_rs.eof
		if l_rs("pestip") = "B" then
			AuxB = l_rs("motfornro")
		end if
		if l_rs("pestip") = "T" then
			AuxT = l_rs("motfornro")
		end if
		l_rs.movenext
	loop
	MotivoForzado = AuxB & "-" & AuxT

l_rs.close

end function


Function MotivoDiferencia(nromov)
' Datos Entrada: Un numero de Movimiento
' Datos Salida: El Codigo del motivo de Diferencia
' Busca los codigo de los Motivos de Diferencia si los hubiera de los pesos de tara y bruto y los separa por un guion

Dim l_rs
Dim l_sql
Dim AuxB
Dim AuxT

Set l_rs = Server.CreateObject("ADODB.RecordSet")
	l_sql = " SELECT pestip,motnro "
	l_sql = l_sql & " FROM tkt_pesaje "
	l_sql = l_sql & " WHERE tkt_pesaje.movnro = " & nromov '& " AND 
	l_sql = l_sql & " "
	rsOpen l_rs, cn, l_sql, 0
	AuxB = ""
	AuxT = ""	
	do while not l_rs.eof
		if l_rs("pestip") = "B" then
			AuxB = l_rs("motnro")
		end if
		if l_rs("pestip") = "T" then
			AuxT = l_rs("motnro")
		end if
		l_rs.movenext
	loop
	MotivoDiferencia = AuxB & "-" & AuxT

l_rs.close

end function


Function Empresa
' Datos Entrada: 
' Datos Salida: Devuelve el nombre de la Empresa configurada en la tabla tkt_config

Dim l_rs
Dim l_sql


Set l_rs = Server.CreateObject("ADODB.RecordSet")
	l_sql = " select empdes "
	l_sql = l_sql & " from tkt_config "
	l_sql = l_sql & "  inner join tkt_empresa on tkt_empresa.empnro = tkt_config.empnro  "
	rsOpen l_rs, cn, l_sql, 0
	if not l_rs.eof then
		Empresa = l_rs(0)
	else 	
		Empresa = ""
	end if
l_rs.close
end function

Function Lugar
' Datos Entrada: 
' Datos Salida: Devuelve el nombre del Lugar configurado en la tabla tkt_config

Dim l_rs
Dim l_sql


Set l_rs = Server.CreateObject("ADODB.RecordSet")
	l_sql = " select lugdes "
	l_sql = l_sql & " from tkt_config "
	l_sql = l_sql & " inner join tkt_lugar on tkt_lugar.lugnro = tkt_config.lugnro  "
	rsOpen l_rs, cn, l_sql, 0	
	if not l_rs.eof then
		Lugar = l_rs(0)
	else 	
		Lugar = ""
	end if
l_rs.close
end function

Function NroInstalacion
' Datos Entrada: 
' Datos Salida: Devuelve el nro de instalación configurado en la tabla tkt_config

Dim l_rs
Dim l_sql


Set l_rs = Server.CreateObject("ADODB.RecordSet")
	l_sql = " select nroins "
	l_sql = l_sql & " from tkt_config "
	rsOpen l_rs, cn, l_sql, 0	
	if not l_rs.eof then
		NroInstalacion = l_rs(0)
	else 	
		NroInstalacion = ""
	end if
l_rs.close
end function

function mostrarEstado(movest)
	Select case movest
		case "T"
			mostrarEstado =  "Terminado"
		case "P"
			mostrarEstado =  "Pendiente"
		case "A"
			mostrarEstado =  "Anulado"
		case else
			mostrarEstado =  l_rs("movest")
	end select
end function
%>
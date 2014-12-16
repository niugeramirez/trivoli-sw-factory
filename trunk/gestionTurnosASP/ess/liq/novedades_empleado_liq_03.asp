<% Option Explicit %>
<!--#include virtual="/serviciolocal/shared/inc/sec.inc"-->
<!--#include virtual="/serviciolocal/shared/inc/const.inc"-->
<!--#include virtual="/serviciolocal/shared/db/conn_db.inc"-->
<!--#include virtual="/serviciolocal/shared/inc/fecha.inc"-->
<!--#include virtual="/serviciolocal/shared/inc/sqls.inc"-->
<!--#include virtual="/serviciolocal/ess/shared/inc/accesoESS.inc"-->
<!--
Archivo: novedades_empleado_liq_03.asp
Descripción: 
Autor: FFavre
Fecha: 23-10-03
Modificado:
	16-11-03 FFavre Se agrego firma 
	24-11-03 FFavre Se agregaron los periodos retroactivos.
 	04-02-04 FFavre Se le cambio el nombre a dos parametros (nepliqdes - nepliqhas) que le son pasados a la ventana.
	12-02-04 F.Favre Actualiza en la ventana llamadora segun sea A o M.
    03-09-04 - Scarpa D. - Usar como clave de la SQL el nenro
    05-10-04 - Scarpa D. - Correccion de las novedades retroactivas	
	25-10-05 - Leticia A. - Agregar Adecuacion a Autogestion - se comento lo de Firmas.
	27-10-05 - Leticia A. - Si el confrep esta config, controlar que en el alta solo se carguen conceptos/param que estan configurados.
-->
<%
 on error goto 0
 
 Dim l_tipo
 Dim l_cm
 Dim l_rs
 Dim l_sql
 
 Dim l_nenro
 Dim l_empleado
 Dim l_concnro
 Dim l_tpanro
 Dim l_nevalor
 Dim l_nevigencia
 Dim l_nedesde
 Dim l_nehasta
 Dim l_neretro
 Dim l_nepliqdesde
 Dim l_nepliqhasta
 Dim l_netexto
 Dim l_cysfirmas
 Dim l_cysfirmas1
 Dim l_concretro
 Dim l_repnro 
 Dim l_sql_confrep
 
 ' ************
 l_repnro = 150
 
 l_empleado = l_ess_ternro
 l_tipo = request.QueryString("tipo")
 
 l_nenro	   = request.Form("nenro")
 'l_empleado	   = request.Form("ternro")
 l_concnro 	   = request.Form("concnro")
 l_tpanro 	   = request.Form("tpanro")
 l_nevalor 	   = request.Form("nevalor")
 l_nevigencia  = request.Form("nevigencia")
 l_nedesde 	   = request.Form("nedesde")
 l_nehasta 	   = request.Form("nehasta")
 l_nepliqdesde = request.Form("nepliqdes")
 l_nepliqhasta = request.Form("nepliqhas")
 l_netexto	   = request.Form("netexto")
 l_concretro   = request.Form("concretro")
 'l_neretro 	  = request.Form("neretro")
 

' ______________________________________________________________
' Verificar si se cargaron Conceptos a mostrar en el ConfRep    
 Set l_rs = Server.CreateObject("ADODB.RecordSet")
 l_sql = " SELECT repnro FROM confrep WHERE repnro=" & l_repnro
 rsOpen l_rs, cn, l_sql, 0 
 
 l_sql_confrep = ""
 if not l_rs.eof then 
 	l_rs.Close
	
	l_sql = " SELECT * FROM confrep "
	l_sql = l_sql & " INNER JOIN concepto ON UPPER(concepto.conccod)=UPPER(confrep.confval2) "
	l_sql = l_sql & " WHERE confrep.repnro="& l_repnro & " AND confrep.confval="& l_tpanro & " AND concepto.concnro="& l_concnro
	rsOpen l_rs, cn, l_sql, 0 
	
	if l_rs.eof then 
		Response.write "<script>" & vbCrLf
		response.write " alert('El Concepto/Parámetro no está habilitado para el ingreso.');"
		Response.write " window.close();" & vbCrLf
		Response.write "</script>" & vbCrLf
		response.end 
	end if
 end if 
 l_rs.Close
 
 set l_rs = Nothing
' ______________________________________________________________
 

'Chequea los checkbox segun su valor y las fechas
 if len(l_nevigencia) > 0 		then  l_nevigencia = "-1" 										else l_nevigencia  = "0" 		end if
 if l_nedesde <> "" 			then  l_nedesde = cambiafecha(l_nedesde, "YMD", true)			else l_nedesde 	   = "null"		end if
 if l_nehasta <> "" 			then  l_nehasta = cambiafecha(l_nehasta, "YMD", true)			else l_nehasta 	   = "null"		end if
 if len(l_concretro) <> 0 then
    if l_nepliqdesde = ""			then  l_nepliqdesde = "null"	end if
    if l_nepliqhasta = ""			then  l_nepliqhasta = "null"	end if
 else
    l_nepliqdesde = "null"
    l_nepliqhasta = "null"
 end if

 l_neretro = "null"
 
 Set l_cm = Server.CreateObject("ADODB.Command")
 
 if l_tipo = "A" then 
	l_sql = "INSERT INTO novemp "
	l_sql = l_sql & "(empleado, concnro, tpanro, nevalor, nevigencia, nedesde, nehasta, neretro, nepliqdesde, nepliqhasta,netexto)"
	l_sql = l_sql & " values (" & l_empleado  & ", " & l_concnro & ", " & l_tpanro & ", " & l_nevalor & ", "
	l_sql = l_sql & l_nevigencia & ", " & l_nedesde & ", " & l_nehasta & ", " & l_neretro & ", " & l_nepliqdesde & ", " & l_nepliqhasta & ",'" & l_netexto &"' ) "
 else
	l_sql = "UPDATE novemp "
	l_sql = l_sql & "set nevalor=" & l_nevalor & ", nevigencia=" & l_nevigencia & ", "
	l_sql = l_sql & "nedesde=" & l_nedesde & ", nehasta=" & l_nehasta & ", neretro=" & l_neretro & ", "
	l_sql = l_sql & "nepliqdesde=" & l_nepliqdesde & ", nepliqhasta=" & l_nepliqhasta & ", "
	l_sql = l_sql & "netexto='" & l_netexto & "' "
	l_sql = l_sql & "WHERE nenro = " & l_nenro
 end if
' response.write l_sql
 
 cn.beginTrans
 l_cm.activeconnection = Cn
 l_cm.CommandText = l_sql
 cmExecute l_cm, l_sql, 0
 
'------------------------------------------------------------------------------------------------------------------
' Firmas
' l_cysfirmas  = request.Form("seleccion")
' l_cysfirmas1 = request.Form("seleccion1")
 
 if l_tipo = "A" then 
	Set l_rs = Server.CreateObject("ADODB.RecordSet")
	l_sql = fsql_seqvalue("nov_id","novemp")
	rsOpen l_rs, cn, l_sql, 0
	l_nenro = l_rs("nov_id")
	l_rs.Close
 end if
 
 if (CStr(l_cysfirmas1) <> "") then
 	'if inStr(l_cysfirmas1,"@@@") <> 0 then
    	'l_cysfirmas1 = left(l_cysfirmas1,inStr(l_cysfirmas1,"@@@")-1) & l_nenro & mid(l_cysfirmas1,inStr(l_cysfirmas1,"@@@")+3)
	'end if
  	'l_cm.activeconnection = Cn
  	'l_cm.CommandText = l_cysfirmas1
  	'cmExecute l_cm, l_cysfirmas1, 0
 end if  
 
 if (CStr(l_cysfirmas) <> "") then
 	'if inStr(l_cysfirmas,"@@@") <> 0 then
    	'l_cysfirmas = left(l_cysfirmas,inStr(l_cysfirmas,"@@@")-1) & l_nenro & mid(l_cysfirmas,inStr(l_cysfirmas,"@@@")+3)
	'end if
  	'l_cm.activeconnection = Cn
  	'l_cm.CommandText = l_cysfirmas
  	'cmExecute l_cm, l_cysfirmas, 0
 end if  
'------------------------------------------------------------------------------------------------------------------
 
 cn.CommitTrans
 
 Response.write "<script>" & vbCrLf
 Response.write "alert('Operación realizada');" & vbCrLf
 'if l_tipo = "A" then
	'Response.write "window.opener.Blanquear();" & vbCrLf  '  para que ?????????
 'else
	Response.write "window.opener.Salir();" & vbCrLf
 'end if
 
 Response.write "window.close();" & vbCrLf
 Response.write "</script>" & vbCrLf
 
 Set cn = Nothing
 Set l_cm = Nothing
 %>
 



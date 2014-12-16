<% Option Explicit %>
<!--#include virtual="/serviciolocal/shared/inc/sec.inc"-->
<!--#include virtual="/serviciolocal/shared/inc/const.inc"-->
<!--#include virtual="/serviciolocal/shared/inc/const_eva.inc"-->
<!--#include virtual="/serviciolocal/shared/db/conn_db.inc"-->
<!--#include virtual="/serviciolocal/shared/inc/fecha.inc"-->
<script src="/serviciolocal/shared/js/fn_windows.js"></script>
<script src="/serviciolocal/shared/js/fn_confirm.js"></script>
<script src="/serviciolocal/shared/js/fn_ayuda.js"></script>
<%
'Archivo	: habilitar_evaluador_eva_ag_00.asp
'Descripción: actualizar campos habilitado y cargado y habilitar el proximo
'Autor		: CCRossi 
'Fecha		: 31-05-2004
'Modificacion	: 04-11-2004 CCRossi- si Termina el evaluador (ABN) entonces copiar
				' los datos al auto y cerrar elauto tambien
'Modificacion	: 04-11-2004 CCRossi- si Termina el SUPER (ABN) entonces 
'			      cierro todas las secciones
'Modificacion	: 08-11-2004 CCRossi- si Termina el SUPER (ABN) entonces 
'			      cierro todas las secciones OBLIGATORIAS UNICAMENTE
'				: 30-03-2005 - LA: - deloitte - CabAprobada - aprobar evalaucion
'				: 16-04-2005 - LA. verificar si es la ultima seccion, habilitar roles s/.
'Integrado.-	
'				: 10-08-2005 - L.A. - tener en cuenta en la ult seccion cdo Auto=Gerente , Rev=Spcio y Rev=Gerente
'-------------------------------------------------------------------------------
on error goto 0

 Dim l_cm
 Dim l_sql
 Dim l_rs
 Dim l_rs1, l_rs2
 Dim l_sql1
 
'variables locales
 dim l_habilitarsiguiente
 dim l_evacabnro 
 dim l_hora 
 dim l_arrhr
 dim l_evatevnro
 dim l_evaluador
 dim l_tipsecobj
 dim l_evaseccmail
 
 dim l_ultimaseccion
 dim l_evaseccnroUlt
 dim l_evldrnroUlt
 dim l_evatevnroUlt 
 dim l_evaluadorUlt
 dim l_evaseccmailUlt
 
 dim l_evaluadorAux 
 dim l_evldrnroAux  
 
 dim l_termino
 dim l_elimino
 dim l_seccioncargada  ' guarda el valor --> de si el ultimo rol termino la seccion (evldorcargada).
 
'parametros de entrada
 Dim l_evldrnro
 Dim l_evldorcargada
 Dim l_evaseccnro
 
'uso local
 dim i
 dim j

 dim l_evaacuerdo
 dim l_cierre
 dim l_rsx
 dim l_proximoevatevnro
   
 l_evldrnro		 = Request.QueryString("evldrnro")
 l_evldorcargada = request.QueryString("evldorcargada")
 l_evaseccnro	 = request.QueryString("evaseccnro")

function strto2(cad)
	if trim(cad) <>"" then
		if len(cad)<2 then
			strto2= "0" & cad
		else
			strto2= cad
		end if 
	else
		strto2= "00"
	end if	
end function

' _________________________________________________________
'  Aprobar cabecera 
sub aprobarCab(evacabnro, cabaprobada)
 
set l_cm = Server.CreateObject("ADODB.Command")
l_sql = "UPDATE  evacab SET "
l_sql = l_sql & " cabaprobada=" & cabaprobada & ","
l_sql = l_sql & " fechaapro = " & cambiafecha(Date(),"","") & ","
l_sql = l_sql & " horaapro  =	'" & l_hora & "'"
l_sql = l_sql & " WHERE evacabnro="& evacabnro
l_cm.activeconnection = Cn
l_cm.CommandText = l_sql
cmExecute l_cm, l_sql, 0

set l_cm = Server.CreateObject("ADODB.Command")
l_sql = "UPDATE  evadetevldor SET "
l_sql = l_sql & " habilitado = 0 "
l_sql = l_sql & " WHERE evacabnro="& l_evacabnro
l_cm.activeconnection = Cn
l_cm.CommandText = l_sql
cmExecute l_cm, l_sql, 0
end sub


'XXXXXXXXXXXXXXXXXXXXXX
' _____________________________________________________________________________________________________________
' si es la ultima seccion --> ver si los roles son iguales --> cargar las otras evluacs.
' Si autoevaluador = Gerente -->  si se cargo Autoev (evldorcargada=-1) entonces cargo la evaluacion de Gerente,
'		 dado que para que Autoev. este habilitado --> todas secc oblig de Autoev. y gerente estan terminadas.  
' ______________________________________________________________________________________________________________
sub cargarEvldorIguales(evacabnro, evaluador, evatevnro, evaseccnroUlt)
dim evldrnroCargar
evldrnroCargar  =  ""

if evatevnro=cautoevaluador or evatevnro=cevaluador then
	
	Set l_rs1 = Server.CreateObject("ADODB.RecordSet")
	 ' proygerente o proyrevisor nova! , xq si el evaluado es el revisor , ent su revisor pasa a ser el gerente..
	l_sql = " SELECT evldrnro "
	l_sql = l_sql & " FROM evaproyecto "
	l_sql = l_sql & " INNER JOIN evacab ON evacab.evaproynro = evaproyecto.evaproynro "
	l_sql = l_sql & " INNER JOIN evadetevldor ON evadetevldor.evacabnro = evacab.evacabnro "
	l_sql = l_sql & " INNER JOIN evaoblieva ON evaoblieva.evaseccnro = evadetevldor.evaseccnro AND evaoblieva.evatevnro = evadetevldor.evatevnro " ' evatevobli=-1
	l_sql = l_sql & " WHERE evacab.evacabnro="& evacabnro & " AND evadetevldor.evaseccnro ="& evaseccnroUlt
	l_sql = l_sql & "   AND evadetevldor.evatevnro <> "& cautoevaluador & " AND evadetevldor.evatevnro <> "& cevaluador 
	l_sql = l_sql & "  AND  evadetevldor.evaluador="& evaluador 'l_evaluador 
	if evatevnro= cautoevaluador then
		l_sql = l_sql & " ORDER BY evaoblieva.evaobliorden "
	else
		l_sql = l_sql & " ORDER BY evaoblieva.evaobliorden DESC " 
	end if 	
	
	rsOpen l_rs1, cn, l_sql, 0
	if not l_rs1.EOF then 
		evldrnroCargar = l_rs1("evldrnro") 
	end if 
	l_rs1.Close 
	set l_rs1=nothing 

	' algun rol es igual a otro --> doy por terminada el evldrnro
	if evldrnroCargar <> "" then	
		set l_cm = Server.CreateObject("ADODB.Command")
		l_sql = "UPDATE  evadetevldor SET  habilitado=0, evldorcargada = -1," 
		l_sql = l_sql & " fechacar ="  & cambiafecha(Date(),"","") & ","
		l_sql = l_sql & " horacar ='"  & l_hora & "'"
		l_sql = l_sql & " WHERE evldrnro = " & evldrnroCargar
		l_cm.activeconnection = Cn
		l_cm.CommandText = l_sql
		cmExecute l_cm, l_sql, 0
	end if
end if  ' de evatevnro=cautoevaluador or evatevnro=cevaluador..

end sub
' XXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXx



' _____________________________________________________________________________________
'    busco los roles que no terminaron de la ultima seccion 						   
' _____________________________________________________________________________________
sub rolesUltSeccion (evaseccnroUlt, evldrnroUlt, evatevnroUlt, evaluadorUlt, evaseccmailUlt)
Set l_rs = Server.CreateObject("ADODB.RecordSet")

l_sql = "SELECT evasecc.evaseccnro, evadetevldor.evldrnro, evadetevldor.evatevnro, evadetevldor.evaluador, evasecc.evaseccmail "
l_sql = l_sql & " FROM  evacab "
l_sql = l_sql & " INNER JOIN evadetevldor ON evadetevldor.evacabnro = evacab.evacabnro "  '  AND evadetevldor.evatevnro="& cautoevaluador
l_sql = l_sql & " INNER JOIN evadet ON evadet.evacabnro = evacab.evacabnro "
l_sql = l_sql & " INNER JOIN evasecc ON evasecc.evaseccnro = evadet.evaseccnro AND evasecc.evaseccnro = evadetevldor.evaseccnro "
l_sql = l_sql & " INNER JOIN evaoblieva ON evaoblieva.evaseccnro= evadetevldor.evaseccnro"
l_sql = l_sql & "		AND  evaoblieva.evatevnro= evadetevldor.evatevnro AND evaoblieva.evatevobli=-1 "
l_sql = l_sql & " WHERE evacab.evacabnro =" & l_evacabnro & " AND evasecc.ultimasecc = -1 and evasecc.evaseccnro="& evaseccnroUlt
l_sql = l_sql & "       AND evadetevldor.evldorcargada=0 "  ' miro las secciones que no  se terminaron
l_sql = l_sql & " ORDER BY evaoblieva.evaobliorden "
rsOpen l_rs, cn, l_sql, 0
response.write " roles ult secc  " & l_sql & "<br><br>"
if not l_rs.EOF then
	'evaseccnroUlt = l_rs("evaseccnro")
	evldrnroUlt    = l_rs("evldrnro")
	evatevnroUlt   = l_rs("evatevnro")
	evaluadorUlt   = l_rs("evaluador")
	evaseccmailUlt = l_rs("evaseccmail")
else
	evldrnroUlt    = "" 
	evatevnroUlt   = ""
	evaluadorUlt   = ""
	evaseccmailUlt = ""
end if
l_rs.Close
set l_rs=nothing
end sub


'_______________________________________________________________________________________
'  buscar todos los evadetevldor que tiene que terminar el Evaluador para un empleado   
'_______________________________________________________________________________________
sub seccsObligsTerminada (tipoevaluador, evaluador, termino) 
dim l_sqltipoev

termino= "NO"

if tipoevaluador="" or isNull(tipoevaluador) then 
		' verrr - 		' puede ser porque se modificaron datos desde el modulo!!
else 
	l_sqltipoev = ""
	response.write " tipoeval. " & tipoevaluador & "<br><br>"
	' ver si tipoevaluador --> auto = gerente, rev=socio o rev =gerente ..... 
	if tipoevaluador = cautoevaluador or tipoevaluador=cevaluador  then  ' ---
		Set l_rs1 = Server.CreateObject("ADODB.RecordSet") 
		l_sql = " SELECT evldrnro, evadetevldor.evatevnro "
		l_sql = l_sql & " FROM evaproyecto "
		l_sql = l_sql & " INNER JOIN evacab ON evacab.evaproynro = evaproyecto.evaproynro "
		l_sql = l_sql & " INNER JOIN evadetevldor ON evadetevldor.evacabnro = evacab.evacabnro "
		l_sql = l_sql & " INNER JOIN evaoblieva ON evaoblieva.evaseccnro = evadetevldor.evaseccnro AND evaoblieva.evatevnro = evadetevldor.evatevnro " ' evatevobli=-1
		l_sql = l_sql & " WHERE evacab.evacabnro="& l_evacabnro & " AND evadetevldor.evaseccnro ="& l_evaseccnroUlt 
		l_sql = l_sql & "   AND evadetevldor.evatevnro <> "& cautoevaluador & " AND evadetevldor.evatevnro <> "& cevaluador 
		l_sql = l_sql & "   AND  evadetevldor.evaluador="& evaluador 
		response.write " evldor igual?  " & l_sql & "<br>"
		if l_evatevnro= cautoevaluador then 
			l_sql = l_sql & " ORDER BY evaoblieva.evaobliorden "
		else 
			l_sql = l_sql & " ORDER BY evaoblieva.evaobliorden ASC " 
		end if 
		
		rsOpen l_rs1, cn, l_sql, 0 
		if not l_rs1.EOF then 
			l_sqltipoev = " evadetevldor.evatevnro="& l_rs1("evatevnro") 
		end if 
		l_rs1.Close 
		set l_rs1=nothing 
		
	end if
	
	' ________________________________________________________________________________
	Set l_rs1 = Server.CreateObject("ADODB.RecordSet") 
	l_sql = "SELECT DISTINCT  evadetevldor.evldrnro  " 
	l_sql = l_sql & " FROM evadetevldor " 
	l_sql = l_sql & " INNER JOIN evacab ON evacab.evacabnro= evadetevldor.evacabnro" 
	l_sql = l_sql & " INNER JOIN evadet ON evadet.evacabnro = evacab.evacabnro " 
	l_sql = l_sql & " INNER JOIN evasecc ON evasecc.evaseccnro= evadetevldor.evaseccnro"
	l_sql = l_sql & "		AND evasecc.evaoblig= -1 AND evasecc.ultimasecc <> -1"
	l_sql = l_sql & " INNER JOIN evaoblieva ON evaoblieva.evaseccnro= evadetevldor.evaseccnro"
	l_sql = l_sql & "		AND  evaoblieva.evatevnro= evadetevldor.evatevnro AND evaoblieva.evatevobli=-1 "
	l_sql = l_sql & " WHERE  evacab.evacabnro=" & l_evacabnro  & " AND evadetevldor.evldorcargada=0 " ' si alguna seccion no se termino --> no habilito Evaluador
	if l_sqltipoev <> "" then
		l_sql = l_sql & "    AND ( evadetevldor.evatevnro="& tipoevaluador & " OR " & l_sqltipoev  &")"
	else
		l_sql = l_sql & "    AND evadetevldor.evatevnro="& tipoevaluador
	end if
	
	rsOpen l_rs1, cn, l_sql, 0 
		response.write " secc oblig ter " & l_sql & "<br><br>" 
	if l_rs1.eof then  ' todas las secc oblig para el Evaluador. se terminaron -> le habilito la ultima seccion
		termino= "SI"
	end if 
	l_rs1.Close 
	set l_rs1 = Nothing 
end if 

end sub 

' _________________________________________________________________
' Habilita al evaluador para que complete la seccion.              
' _________________________________________________________________
sub habilitarEvadetevldor(evldrnro, evaseccmail)
	
Set l_rs1 = Server.CreateObject("ADODB.RecordSet")
l_sql = "SELECT habilitado FROM  evadetevldor WHERE evldrnro ="& evldrnro
rsOpen l_rs1, cn, l_sql, 0

if not l_rs1.EOF then 
	if l_rs1("habilitado") <> -1 then
		l_hora = mid(time,1,8)
		l_arrhr= Split(l_hora,":")
		l_hora = strto2(l_arrhr(0))&l_arrhr(1)
		
		set l_cm = Server.CreateObject("ADODB.Command")
		l_sql = "UPDATE  evadetevldor SET " 
		l_sql = l_sql & " habilitado	=  -1,"  
		l_sql = l_sql & " fechahab =  "  & cambiafecha(Date(),"","") & ","
		l_sql = l_sql & " horahab =   '"  & l_hora & "'"
		l_sql = l_sql & " WHERE evldrnro = " & evldrnro
		l_cm.activeconnection = Cn
		l_cm.CommandText = l_sql
		cmExecute l_cm, l_sql, 0
		
		if cUsaMail= -1 and  evaseccmail = -1 then %>
			<script>
				abrirVentanaH('enviomail_eva_00.asp?evldrnro=<%=evldrnro%>', '','','');
			</script>
<%		end if 
		
	end if
end if
l_rs1.Close
set l_rs1=nothing

end sub



' ------------------------------------------------------------------------------------
'											BODY                 					  
' ------------------------------------------------------------------------------------


'Actualizar el evadetevldor actual ________________________________________
'cn.begintrans
l_hora = mid(time,1,8)
l_arrhr= Split(l_hora,":")
l_hora = strto2(l_arrhr(0))&l_arrhr(1)

set l_cm = Server.CreateObject("ADODB.Command")
l_sql = "UPDATE  evadetevldor SET " 
l_sql = l_sql & " habilitado	=  0, evldorcargada = -1," 
l_sql = l_sql & " fechacar		=    "  & cambiafecha(Date(),"","") & "," 
l_sql = l_sql & " horacar		=   '"  & l_hora & "'" 
l_sql = l_sql & " WHERE evldrnro = " & l_evldrnro 
l_cm.activeconnection = Cn 
l_cm.CommandText = l_sql 
cmExecute l_cm, l_sql, 0 
'_________________________________________________________________________


Set l_rs = Server.CreateObject("ADODB.RecordSet")
l_sql = "SELECT evacabnro, evatevnro, evaluador FROM  evadetevldor "
l_sql = l_sql & " WHERE evldrnro = " & l_evldrnro
rsOpen l_rs, cn, l_sql, 0
if not l_rs.EOF then
	l_evacabnro = l_rs("evacabnro")
	l_evatevnro = l_rs("evatevnro")
	l_evaluador = l_rs("evaluador")
end if
l_rs.Close
set l_rs=nothing


l_ultimaseccion="NO"


if cint(cdeloitte) = -1 then 
	l_evaseccnroUlt = "" 
	Set l_rs = Server.CreateObject("ADODB.RecordSet") 
	l_sql = "SELECT evasecc.evaseccnro FROM  evacab " 
	l_sql = l_sql & " INNER JOIN evadetevldor ON evadetevldor.evacabnro = evacab.evacabnro "
	l_sql = l_sql & " INNER JOIN evasecc ON evasecc.evaseccnro = evadetevldor.evaseccnro" 
	l_sql = l_sql & " WHERE evasecc.ultimasecc = -1 and evacab.evacabnro="& l_evacabnro  
	rsOpen l_rs, cn, l_sql, 0
	if not l_rs.EOF then 
		l_evaseccnroUlt = l_rs("evaseccnro") 
	end if 
	l_rs.Close 
	set l_rs=nothing 
	
	' _________________________________________________
	if trim (l_evaseccnroUlt)= trim(l_evaseccnro) then 
		l_ultimaseccion="SI" 
		cargarEvldorIguales  l_evacabnro,l_evaluador, l_evatevnro,l_evaseccnroUlt  
	end if 
	
	
	'  busco los roles que no terminaron de la ultima seccion ________________________
	rolesUltSeccion  l_evaseccnroUlt, l_evldrnroUlt, l_evatevnroUlt, l_evaluadorUlt, l_evaseccmailUlt 
	
	
	if l_ultimaseccion="SI" then 
		
		'Set l_rs = Server.CreateObject("ADODB.RecordSet")
		'l_sql = "SELECT evaoblieva.evatevnro, evaobliorden, afteranterior, evldrnro, habilitado, evaseccmail "
		'l_sql = l_sql & " FROM  evaoblieva "
		'l_sql = l_sql & " INNER JOIN evadetevldor on evadetevldor.evatevnro = evaoblieva.evatevnro " 
		'l_sql = l_sql & "   AND  evadetevldor.evaseccnro = evaoblieva.evaseccnro "
		'l_sql = l_sql & " INNER JOIN evasecc on evasecc.evaseccnro = evaoblieva.evaseccnro " 
		'l_sql = l_sql & " WHERE evaoblieva.evaseccnro= " & l_evaseccnro
		'l_sql = l_sql & "   AND   evacabnro= " & l_evacabnro & " AND evldorcargada= 0" 
		'l_sql = l_sql & " ORDER BY evaobliorden	"
		'rsOpen l_rs, cn, l_sql, 0
		
		'if not l_rs.eof then 
			'seccsObligsTerminada l_rs("evatevnro"), l_termino 
			'if l_termino= "SI" then  ' todas las secc oblig para el tipoev. se terminaron -> le habilito la ultima seccion
				'habilitarEvadetevldor l_rs("evldrnro"), l_rs("evaseccmail")
			'end if 
		
		if l_evatevnroUlt <> "" then ' preg por el rol de la ult secc. 
			seccsObligsTerminada l_evatevnroUlt, l_evaluadorUlt, l_termino ' ---l_proxRol,
			if l_termino= "SI" then  
				' todas las secc oblig para el tipoev. se terminaron -> le habilito la ultima seccion
				habilitarEvadetevldor l_evldrnroUlt, l_evaseccmailUlt   ' l_evaseccmail 
			end if 
			
		else  ' todas secciones terminadas --> apruebo cabecera (evaluacion)
			aprobarCab l_evacabnro, -1  
%>
			<script>
				//abrirVentana('mailavisofineval_eva_00.asp?evldrnro=<%=l_evldrnro%>', '','dialogWidth:200;dialogHeight:200');
				abrirVentanaH('mailavisofineval_eva_00.asp?evldrnro=<%=l_evldrnro%>', '','','');
			</script>
<%		end if
	
		'l_rs.Close
		'set l_rs=nothing
	end if ' ultima seccion 
end if ' deloitte=-1



if l_ultimaseccion = "NO" then 
			
	'BUSCAR LOS EVADETEVLADOR POR ORDEN DE EVATEV 
	Set l_rs = Server.CreateObject("ADODB.RecordSet")
	l_sql = "SELECT evaoblieva.evatevnro, evaobliorden, afteranterior, evldrnro, habilitado, evldorcargada, evaseccmail "
	l_sql = l_sql & " FROM  evaoblieva "
	l_sql = l_sql & " INNER JOIN evadetevldor on evadetevldor.evatevnro = evaoblieva.evatevnro " 
	l_sql = l_sql & "   and  evadetevldor.evaseccnro = evaoblieva.evaseccnro "
	l_sql = l_sql & " INNER JOIN evasecc on evasecc.evaseccnro = evaoblieva.evaseccnro " 
	l_sql = l_sql & " WHERE evaoblieva.evaseccnro= " & l_evaseccnro
	l_sql = l_sql & " AND   evacabnro= " & l_evacabnro
	l_sql = l_sql & " ORDER BY evaobliorden	"
	rsOpen l_rs, cn, l_sql, 0

	l_habilitarsiguiente=0 
	l_seccioncargada = 0
	
	do while not l_rs.eof
		l_evaseccmail = l_rs("evaseccmail")
		l_proximoevatevnro=l_rs("evatevnro") ' se usa para Codelco
		
		'Response.Write("<script>alert('"&l_habilitarsiguiente&"');</script>")
		if l_rs("afteranterior")=0 then 
			'si NO depende del anterior pero ya esta cargado lo inhabilito, y habilito siguiente
			   if l_rs("evldorcargada")	=-1 then
					set l_cm = Server.CreateObject("ADODB.Command")
					l_sql = "UPDATE  evadetevldor SET "
					l_sql = l_sql & " habilitado = 0 " 
					l_sql = l_sql & " WHERE evldrnro = " & l_rs("evldrnro")
					l_sql = l_sql & " AND   evldrnro <> " & l_evldrnro
					l_cm.activeconnection = Cn
					l_cm.CommandText = l_sql
					cmExecute l_cm, l_sql, 0
					
					l_habilitarsiguiente = -1
				else	
					l_hora = mid(time,1,8)
					l_arrhr= Split(l_hora,":")
					l_hora = strto2(l_arrhr(0))&l_arrhr(1)
					set l_cm = Server.CreateObject("ADODB.Command")
					l_sql = "UPDATE  evadetevldor SET "
					l_sql = l_sql & " habilitado =  -1, "  
					l_sql = l_sql & " fechahab =  "  & cambiafecha(Date(),"","") & ","
					l_sql = l_sql & " horahab =   '"  & l_hora & "'"
					l_sql = l_sql & " WHERE evldrnro = " & l_rs("evldrnro")
					l_sql = l_sql & " AND   evldrnro <> " & l_evldrnro
					l_cm.activeconnection = Cn
					l_cm.CommandText = l_sql
					cmExecute l_cm, l_sql, 0
				
					if cUsaMail= -1 and l_evaseccmail = -1 then%>
					<script>
					abrirVentanaH('enviomail_eva_00.asp?evldrnro=<%=l_rs("evldrnro")%>', '','','');
					</script>
					<%end if

					l_habilitarsiguiente = 0
				end if
		else
				' si  depende del anterior y el anterior lo habilito... habilito.
			
				if l_habilitarsiguiente = -1  then
			
					if l_rs("evldorcargada") = -1 then ' SI YA ESTA TERMINADO...
						l_habilitarsiguiente = -1
					else	
						l_hora = mid(time,1,8)
						l_arrhr= Split(l_hora,":")
						l_hora = strto2(l_arrhr(0))&l_arrhr(1)
	
						'****************************************************************
						'verificar si es cierre...
						' PARA CODELCO!
						if l_proximoevatevnro=cgarante then
							l_evaacuerdo=-1
							l_cierre="SI"
							
							Set l_rsx = Server.CreateObject("ADODB.RecordSet")
							l_sql = "SELECT evaacuerdo FROM  evacierre"
							l_sql = l_sql & " INNER JOIN evadetevldor on evadetevldor.evldrnro = evacierre.evldrnro " 
							l_sql = l_sql & "   and  evadetevldor.evaseccnro =  " & l_evaseccnro
							l_sql = l_sql & "   AND   evacabnro= " & l_evacabnro
							l_sql = l_sql & " WHERE evacierre.evaetapa=3"   
							rsOpen l_rsx, cn, l_sql, 0
							if not l_rsx.eof then
								l_evaacuerdo=l_rsx("evaacuerdo")
							else
								l_cierre="NO"
							end if
							l_rsx.close
							set l_rsx=nothing
							
							if l_cierre="NO" or (l_cierre="SI" and l_evaacuerdo=0) then 
								set l_cm = Server.CreateObject("ADODB.Command")
								l_sql = "UPDATE  evadetevldor SET "
								l_sql = l_sql & " habilitado =  -1 ,"  
								l_sql = l_sql & " fechahab =  "  & cambiafecha(Date(),"","") & ","
								l_sql = l_sql & " horahab =   '"  & l_hora & "'"
								l_sql = l_sql & " WHERE evldrnro = " & l_rs("evldrnro")
								l_sql = l_sql & " AND   evldrnro <> " & l_evldrnro
								l_cm.activeconnection = Cn
								l_cm.CommandText = l_sql
								cmExecute l_cm, l_sql, 0
								
								if cUsaMail= -1 and l_evaseccmail=-1 then%>
									<script>
									abrirVentanaH('enviomail_eva_00.asp?evldrnro=<%=l_rs("evldrnro")%>', '','','');
									</script>
								<%end if
								l_habilitarsiguiente = 0 
							end if
						else 
							' en cualquier otro evaluador y para los no-codelcos
							set l_cm = Server.CreateObject("ADODB.Command")
							l_sql = "UPDATE  evadetevldor SET "
							l_sql = l_sql & " habilitado =  -1 ,"  
							l_sql = l_sql & " fechahab =  "  & cambiafecha(Date(),"","") & ","
							l_sql = l_sql & " horahab =   '"  & l_hora & "'"
							l_sql = l_sql & " WHERE evldrnro = " & l_rs("evldrnro")
							l_sql = l_sql & " AND   evldrnro <> " & l_evldrnro
							l_cm.activeconnection = Cn
							l_cm.CommandText = l_sql
							cmExecute l_cm, l_sql, 0
						
							if cUsaMail= -1 and l_evaseccmail=-1 then%>
							<script>
							abrirVentanaH('enviomail_eva_00.asp?evldrnro=<%=l_rs("evldrnro")%>', '','','');
							</script>
							<%end if
							l_habilitarsiguiente = 0 
						end if
						'*****************************************************************
					end if	
				else
					set l_cm = Server.CreateObject("ADODB.Command")
					l_sql = "UPDATE  evadetevldor SET "
					l_sql = l_sql & " habilitado =  0 "  
					l_sql = l_sql & " WHERE evldrnro = " & l_rs("evldrnro")
					l_sql = l_sql & " AND   evldrnro <> " & l_evldrnro
					l_cm.activeconnection = Cn
					l_cm.CommandText = l_sql
					cmExecute l_cm, l_sql, 0
				
					l_habilitarsiguiente = 0
				end if
		end if	
			
		l_seccioncargada = l_rs("evldorcargada")
		l_rs.MoveNext
	loop		
	l_rs.close
	set l_rs = nothing 

	' Si el que Termina es el EVALUADOR y es ABN y es COACH 
	' Copia datos al AUTO si es ABN y cerrar AUTO
	if l_evatevnro = cevaluador and cejemplo=-1 then
		' Busco si la seccion es de Objetivos
		Set l_rs = Server.CreateObject("ADODB.RecordSet")
		l_sql = "SELECT tipsecobj "
		l_sql = l_sql & " FROM  evadetevldor "
		l_sql = l_sql & " INNER JOIN evasecc ON evasecc.evaseccnro= evadetevldor.evaseccnro"
		l_sql = l_sql & " INNER JOIN evatiposecc ON evatiposecc.tipsecnro= evasecc.tipsecnro"
		l_sql = l_sql & " WHERE evldrnro = " & l_evldrnro
		rsOpen l_rs, cn, l_sql, 0
		if not l_rs.EOF then
			l_tipsecobj = l_rs("tipsecobj")
		end if
		l_rs.Close
		set l_rs=nothing
		' si la seccion es de OBJETIVOS
		if l_tipsecobj=-1 then
			Set l_rs = Server.CreateObject("ADODB.RecordSet")
			l_sql = "SELECT evaluaobj.evaobjnro,  evaobjalcanza "
			l_sql = l_sql & " FROM  evaluaobj "
			l_sql = l_sql & " WHERE evldrnro = " & l_evldrnro
			l_sql = l_sql & "   AND evaluaobj.evaborrador = 0 "
			rsOpen l_rs, cn, l_sql, 0
			if not l_rs.eof then
				set l_cm = Server.CreateObject("ADODB.Command")
				l_sql = "UPDATE  evadetevldor SET "
				l_sql = l_sql & " habilitado	=  0,"  
				l_sql = l_sql & " evldorcargada = -1,"  
				l_sql = l_sql & " fechacar		=    "  & cambiafecha(Date(),"","") & ","
				l_sql = l_sql & " horacar		=   '"  & l_hora & "'"
				l_sql = l_sql & " WHERE evldrnro   <> " & l_evldrnro
				l_sql = l_sql & " AND   evacabnro  =  " & l_evacabnro
				l_sql = l_sql & " AND   evaseccnro =  " & l_evaseccnro
				l_cm.activeconnection = Cn
				l_cm.CommandText = l_sql
				cmExecute l_cm, l_sql, 0
			end if
			do while not l_rs.EOF 
				if l_rs("evaobjalcanza") <>"" and not isnull(l_rs("evaobjalcanza")) then
				set l_cm = Server.CreateObject("ADODB.Command")
				l_sql = "UPDATE  evaluaobj SET "
				l_sql = l_sql & " evaobjalcanza	  =  " & l_rs("evaobjalcanza") 
				l_sql = l_sql & " WHERE evldrnro  <> " & l_evldrnro
				l_sql = l_sql & " AND   evaobjnro =  " & l_rs("evaobjnro")
				l_sql = l_sql & " AND   evaluaobj.evaborrador = 0 "
				l_cm.activeconnection = Cn
				l_cm.CommandText = l_sql
				cmExecute l_cm, l_sql, 0
				end if
				l_rs.MoveNext
			loop
			l_rs.Close
			set l_rs=nothing
		end if ' es objetivo
	end if ' es Evaluador y es ABN

	' si es el SUPER y es ABN
	' doy por terminada todas las secciones del resto de los evaluadores
	if l_evatevnro <> cevaluador and l_evatevnro <> cautoevaluador and cejemplo=-1 then
		set l_cm = Server.CreateObject("ADODB.Command")
		l_sql = "UPDATE  evadetevldor SET "
		l_sql = l_sql & " habilitado	=  0,"  
		l_sql = l_sql & " evldorcargada = -1,"  
		l_sql = l_sql & " fechacar		=    "  & cambiafecha(Date(),"","") & ","
		l_sql = l_sql & " horacar		=   '"  & l_hora & "'"
		l_sql = l_sql & " WHERE evldrnro   <> " & l_evldrnro
		l_sql = l_sql & " AND   evacabnro  =  " & l_evacabnro
		l_sql = l_sql & " AND   EXISTS "
		l_sql = l_sql & " (SELECT * FROM evasecc WHERE evasecc.evaseccnro=evadetevldor.evaseccnro AND evasecc.evaoblig=-1)"
		l_cm.activeconnection = Cn
		l_cm.CommandText = l_sql
		cmExecute l_cm, l_sql, 0
	end if
	'cn.committrans
	
	if cdeloitte = -1 then   ' preg por la ultima seccion.
		seccsObligsTerminada l_evatevnroUlt, l_evaluadorUlt, l_termino  
		
		if l_termino= "SI" then  ' todas las secc oblig para el Rol se terminaron -> le habilito la ultima seccion
			habilitarEvadetevldor l_evldrnroUlt, l_evaseccmailUlt 
		end if 
	end if
	
end if  ' ultimaseccion="NO"


' si el rol habilitado era el ultimo de la seccion 
' 			-->  aviso al 1º rol de la prox seccion que dependa de esta (depende de evaseccnro)
'  			--> (le avisoa una sola seccion - y al primer rol) --> FALTA Generalizar!!!
if cint(cdeloitte) = -1 then 
	if l_seccioncargada = -1 then  ' el ultimo rol termino la seccion -> miro la sihay secc dependiente
		Set l_rs1 = Server.CreateObject("ADODB.RecordSet")
		l_sql = " SELECT evasecc.evaseccnro, evadetevldor.evldrnro, evasecc.evaseccmail "
		l_sql = l_sql & " FROM evasecc "
		l_sql = l_sql & " INNER JOIN evaoblieva ON evasecc.evaseccnro = evaoblieva.evaseccnro "
		l_sql = l_sql & " INNER JOIN evadetevldor ON evadetevldor.evaseccnro = evasecc.evaseccnro "
		l_Sql = l_sql & " INNER JOIN evacab ON evacab.evacabnro=evadetevldor.evacabnro "
		l_sql = l_sql & " WHERE dependiente= -1 AND dependesecnro="& l_evaseccnro 
		l_sql = l_sql & "  AND evadetevldor.evatevnro = evaoblieva.evatevnro  AND evacab.evacabnro=" & l_evacabnro
		l_sql = l_sql & " ORDER BY evasecc.orden, evaoblieva.evaobliorden "
		rsOpen l_rs1, cn, l_sql, 0
		
		if not l_rs1.EOF then
			if cUsaMail= -1 and l_rs1("evaseccmail")=-1 then  %>
				<script>
				  abrirVentanaH('enviomail_eva_00.asp?evldrnro=<%=l_rs1("evldrnro")%>', '','','');
				</script>
<%			end if
		end if 
		l_rs1.Close 
		set l_rs1=nothing 
		
		
	end if
end if 'deloitte

cn.close
Set cn = Nothing


'response.write "<script>window.returnValue='0';</script>"
response.write "<script>window.returnValue='0';window.close();</script>"

%>
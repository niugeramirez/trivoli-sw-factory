<% Option Explicit %>
<!--#include virtual="/serviciolocal/ess/shared/inc/const.inc"-->
<!--#include virtual="/serviciolocal/shared/db/conn_db.inc"-->
<!--#include virtual="/serviciolocal/shared/inc/fecha.inc"-->
<!--#include virtual="/serviciolocal/ess/shared/inc/accesoESS.inc"-->
<!--
Archivo        : rep_borr_detallado_liq_01.asp
Descripción    : Reporte - borrador detallado 
Autor          : Scarpa D.
Fecha Creacion : 01/06/2004
Modificado     : 
   				15/07/2004 - Scarpa D. - Correccion en las conversiones
   				21/10/2005 - Leticia A. - Adecuarlo para Autogestion.
				22/03/2006 - Mariano - Se eliminaron los AS en las query para que pase con oracle
-->

<% 
on error goto 0

Const l_Max_Lineas_X_Pag = 55
const l_nro_col = 4

Dim l_rs
Dim l_rs2

Dim l_sql

Dim l_nrolinea
Dim l_nropagina

Dim l_encabezado
Dim l_corte
Dim l_cambioEmp
Dim l_conc_detdom

Dim linea

'Parametros

 Dim l_orden
Dim l_fechaestr 
 
 Dim l_empnombre
 Dim l_empleg
 Dim l_cantColumnas
 Dim l_ternro
 Dim l_posicion

 Dim l_activo 
 Dim l_terape 
 Dim l_ternom 
 Dim l_empfoto
 Dim l_empinterno
 Dim l_empemail
 Dim l_nacionaldes
 Dim l_paisdesc
 Dim l_estcivdesabr
 Dim l_tercasape
 Dim l_terfecestciv
 Dim l_terfecing
 Dim l_terfecnac
 Dim l_tersex
 Dim l_telparticular
 Dim l_telcelular

 Dim l_nombre ' guarda ape y nombre del empleado
 dim l_aux

 Dim l_tipimdire
 Dim l_tipimanchodef
 Dim l_tipimaltodef
 Dim l_terimnombre

 dim l_fecha

 Dim l_terape2
 Dim l_ternom2
 Dim l_empvivpropia
 Dim l_emptarinsalubre

 Dim l_nrodoc
 Dim l_tidsigla
 Dim l_cuil
 Dim l_cuilDesc
 Dim l_nrocat

 Dim l_repempleg
 Dim l_repterape
 Dim l_repternom

 Dim l_estado
 Dim l_edad
 Dim l_fecfincont
 Dim l_fecalta
 Dim l_fecbaja
 
l_fechaestr 	= request("fechaestr")
l_posicion		= request("posicion")

l_ternro = l_ess_ternro

'-------------------------------------------------------
Dim l_organizacion

 l_organizacion = request("organizacion")
 
 if l_organizacion = "" then
    l_organizacion = "0"
 end if
 
'------------------------------------------------------- 
Dim l_fases

 l_fases = request("fases")
 
 if l_fases = "" then
    l_fases = "0"
 end if
 
'------------------------------------------------------- 
Dim l_documentos

 l_documentos = request("documentos")
 
 if l_documentos = "" then
    l_documentos = "0"
 end if
 
'-------------------------------------------------------
Dim l_domicilios
Dim l_tipodom

 l_domicilios = request("domicilios")
 
 if l_domicilios = "" then
    l_domicilios = "0"
 else
    l_tipodom = request("tipodom")
	if l_tipodom = "" then
	   l_tipodom = "0"
	end if
 end if
 
'-------------------------------------------------------
Dim l_familiares
Dim l_parentesco

 l_familiares = request("familiares")
 
 if l_familiares = "" then
    l_familiares = "0"
 else
    l_parentesco = request("parentesco")
	if l_parentesco = "" then
	   l_parentesco = "0"
	end if
 end if 

' Imprime el encabezado de cada pagina
sub encabezado
	l_nrolinea = l_nrolinea+1
%>
	<tr> 
	  <td align="center" colspan="<%=l_cantColumnas%>">
	    <table>
		 <tr>
			<td align="center" width="90%" >
				<b>DATOS DEL EMPLEADO </b> <br> &nbsp;
			</td>
	       	<td align="right" width="10%" nowrap>  
				P&aacute;gina: <%= l_posicion & "-" &l_nropagina%> 
			</td>				
		 </tr>	
		 <!--	
		 <tr>
		    <td colspan="2" nowrap>
			   <b>Empleado:&nbsp;</b><%'= l_empleg & " - " &l_empnombre%>
			</td>
		 </tr>
		 -->		 
		</table>		
	  </td>			
	</tr>
<%
end sub 'encabezado

'-----------------------------------------------------------------------------------------------------------
function edad(numero)
	dim l_anos
	l_anos = datediff("yyyy",numero, now)
	if month(numero) >= month(now) then
		if day(numero) > day(now) then
			l_anos = Cint(l_anos) - 1
		end if
	end if
	edad = l_anos
end function

'-----------------------------------------------------------------------------------------------------------
sub titulo_empleado()
%>
  <tr>
    <td class="th2" colspan="<%= l_cantColumnas %>">
	   Datos B&aacute;sicos
	</td>  
  </tr>
  <tr>
    <td colspan="<%= l_cantColumnas %>">
	   <table width="100%" cellpadding="1" cellspacing="1">
<%
   l_nrolinea  = l_nrolinea  + 1

end sub 'titulo_empleado()

'-----------------------------------------------------------------------------------------------------------
sub cerrar_tabla()
%>
	   </table>
	</td>
  </tr>
<%
end sub 'cerrar_tabla()


'-----------------------------------------------------------------------------------------------------------
sub titulo_organizacion()
%>
  <tr>
    <td colspan="<%= l_cantColumnas %>" class="th2">
	   Organizaci&oacute;n
	</td>  
  </tr>
  <tr>
    <td colspan="<%= l_cantColumnas %>">
	   <table width="100%" cellpadding="2" cellspacing="2">
		  <tr>
		     <th class="stytit01">Tipo Estructura</th>
		     <th class="stytit01">Estructura</th>
		     <th class="stytit01">Clase</th>
		     <th class="stytit01">Desde</th>
		     <th class="stytit01">Hasta</th>
		  </tr>
<%

  l_nrolinea  = l_nrolinea  + 2
  
end sub 'titulo_organizacion()


'-----------------------------------------------------------------------------------------------------------
sub titulo_fases()
%>
  <tr>
    <td class="th2" colspan="<%= l_cantColumnas %>">
	   Fases
	</td>  
  </tr>
  <tr>
    <td colspan="<%= l_cantColumnas %>">
	   <table width="100%" cellpadding="1" cellspacing="1">
		  <tr>
		     <th class="stytit01">Fecha Alta</th>
		     <th class="stytit01">Fecha Baja</th>
		     <th class="stytit01">Causa</th>
		     <th class="stytit01">Estado</th>
		     <th class="stytit01">Empleo Anterior</th>
		  </tr>
<%
  l_nrolinea  = l_nrolinea  + 2
  
end sub 'titulo_fases()

'-----------------------------------------------------------------------------------------------------------
sub titulo_documentos()
%>
  <tr>
    <td class="th2" colspan="<%= l_cantColumnas %>">
	   Documentos
	</td>  
  </tr>
  <tr>
    <td colspan="<%= l_cantColumnas %>">
	   <table width="100%" cellpadding="1" cellspacing="1" border="0">
		  <tr>
		     <th class="stytit01">Nombre</th>
		     <th class="stytit01">Sigla</th>
		     <th class="stytit01">N&uacute;mero</th>
		     <th class="stytit01">Vencimento</th>
		  </tr>
<%
  l_nrolinea  = l_nrolinea  + 2
  
end sub 'titulo_documentos()


'-----------------------------------------------------------------------------------------------------------
sub titulo_domicilios()
%>
  <tr>
    <td class="th2" colspan="<%= l_cantColumnas %>">
	   Domicilios
	</td>  
  </tr>
  <tr>
    <td colspan="<%= l_cantColumnas %>">
	   <table width="100%" cellpadding="1" cellspacing="1" border="0">
		  <tr>
		     <th class="stytit01" align="center">Principal</th>
		     <th class="stytit01" align="left">Tipo&nbsp;Domicilio</th>
		     <th class="stytit01" align="left">Calle/N&uacute;mero/Piso</th>
		     <th class="stytit01" align="left">Ofic/Depto</th>
		     <th class="stytit01" align="left">Localidad</th>
		     <th class="stytit01" align="left">CP</th>
		  </tr>
<%

  l_nrolinea  = l_nrolinea  + 2
  
end sub 'titulo_domicilios()


'-----------------------------------------------------------------------------------------------------------
sub titulo_familiares()
%>
  <tr>
    <td class="th2" colspan="<%= l_cantColumnas %>">
	   Familiares
	</td>  
  </tr>
  <tr>
    <td colspan="<%= l_cantColumnas %>">
	   <table width="100%" cellpadding="1" cellspacing="1" border="0">
		  <tr>
		    <th class="stytit01">Parentesco</th>
		    <th class="stytit01">Apellido&nbsp;y&nbsp;Nombre</th>
			<th class="stytit01">Documento</th>
			<th class="stytit01">Fecha&nbsp;Nac.</th>
			<th class="stytit01">Edad</th>
			<th class="stytit01">Estudia</th>
			<th class="stytit01">Estado</th>
		    <th class="stytit01">Salario</th>		
		    <th class="stytit01">DDJJ</th>		
		    <th class="stytit01">Fec.&nbsp;Desde&nbsp;-&nbsp;Hasta</th>		
		  </tr>
<%

  l_nrolinea  = l_nrolinea  + 2
  
end sub 'titulo_familiares()

'-----------------------------------------------------------------------------------------------------------
'Imprime los datos basicos del empleado
sub datos_empleado()

	l_sql = "SELECT empleado.terape, empleado.ternom, empleado.ternro, empleado.empinterno "
	l_sql = l_sql & ", empleado.empemail, nacionaldes, paisdesc, estcivdesabr, tercasape, terfecing, terfecnac "
	l_sql = l_sql & ", tersex, empleado.empnro, empleado.terape2, empleado.ternom2 "
	l_sql = l_sql & ", terfecestciv, nacionaldes, docu.nrodoc, tipodocu.tidsigla, empleado.empfbajaprev "
	l_sql = l_sql & ", cuil.nrodoc as nrocuil, cat.nrodoc nrocat, empleado.empleg repempleg  "
	l_sql = l_sql & ", empleado.empleg repempleg, empleado.terape repterape, empleado.ternom repternom, empleado.empest "
	l_sql = l_sql & "FROM empleado INNER JOIN tercero ON empleado.ternro= tercero.ternro "
	l_sql = l_sql & "LEFT JOIN nacionalidad ON nacionalidad.nacionalnro= tercero.nacionalnro "
	l_sql = l_sql & "LEFT JOIN pais ON pais.paisnro= tercero.paisnro "
	l_sql = l_sql & "LEFT JOIN estcivil ON estcivil.estcivnro= tercero.estcivnro "
	l_sql = l_sql & "LEFT JOIN ter_doc docu ON docu.ternro= tercero.ternro and docu.tidnro>0 and docu.tidnro<5 "
	l_sql = l_sql & "LEFT JOIN tipodocu     ON tipodocu.tidnro= docu.tidnro "
	l_sql = l_sql & "LEFT JOIN ter_doc cuil ON cuil.ternro= tercero.ternro and cuil.tidnro=10 "
	l_sql = l_sql & "LEFT JOIN ter_doc cat  ON cat.ternro= tercero.ternro  and cat.tidnro=20 "
	l_sql = l_sql & "LEFT JOIN empleado empreporta ON empreporta.empreporta = empleado.ternro "
	l_sql = l_sql & "WHERE empleado.empleg=" & l_empleg 
	
	rsOpen l_rs, cn, l_sql, 0
	'Response.Write l_sql
	if not l_rs.EOF then
  	    l_activo = (CInt(l_rs("empest")) = -1)
		l_terape = l_rs("terape")
		l_ternom = l_rs("ternom")
		l_ternro = l_rs("ternro")
		
		l_empinterno = l_rs("empinterno")
		l_empemail = l_rs("empemail")
		l_nacionaldes = l_rs("nacionaldes")
		l_paisdesc = l_rs("paisdesc")
		l_estcivdesabr = l_rs("estcivdesabr")
		l_tercasape = l_rs("tercasape")
		l_terfecestciv = l_rs("terfecestciv")
		l_terfecing = l_rs("terfecing")
		l_terfecnac = l_rs("terfecnac")
		
		if isNull(l_terfecnac) then
	  	    l_edad = ""
		else
			if (month(date()) > month(CDate(l_terfecnac))) then
			   l_edad = DateDiff("yyyy",CDate(l_terfecnac),date()) 	
			else	
			   if   (month(date) = month(CDate(l_terfecnac)))  AND (day(date) > day(CDate(l_terfecnac)))  then
			     l_edad = DateDiff("yyyy",CDate(l_terfecnac),date()) 
			   else
			     l_edad = DateDiff("yyyy",CDate(l_terfecnac),date()) -1
			   end if
			end if
		end if
	
		l_tersex = l_rs("tersex")
		
		l_terape2 = l_rs("terape2")
		l_ternom2 = l_rs("ternom2")
		l_terfecestciv = l_rs("terfecestciv")
			
		l_fecfincont = l_rs("empfbajaprev")
	
		l_nrodoc = l_rs("nrodoc")
		l_tidsigla = l_rs("tidsigla")
		l_cuil   = l_rs("nrocuil")
		l_nrocat = l_rs("nrocat")
		l_repempleg= l_rs("repempleg")
		l_repterape= l_rs("repterape")
		l_repternom= l_rs("repternom")
		
		l_nombre = l_terape 
		if trim(l_terape2)<>"" then
			l_nombre = l_nombre & " "& l_terape2 
		end if
		if trim(l_ternom)<>"" or trim(l_ternom2)<>"" then
			l_nombre = l_nombre & ", " & l_ternom 
		end if
		if trim(l_ternom2)<>"" then
			l_nombre = l_nombre & ", " & l_ternom2 
		end if
	end if	
	l_rs.Close
	
	'Busco el telefono particular y el celular
	l_sql =         " SELECT detdom.domnro,telnro,teldefault, telcelular "
	l_sql = l_sql & " FROM detdom  "
	l_sql = l_sql & " INNER JOIN cabdom ON detdom.domnro=cabdom.domnro AND domdefault = -1 "
	l_sql = l_sql & " INNER JOIN telefono ON telefono.domnro = detdom.domnro AND (teldefault = -1 OR telcelular = -1)"
	l_sql = l_sql & " WHERE cabdom.ternro= " & l_ternro
	
	rsOpen l_rs, cn, l_sql, 0
	
	l_telparticular = ""
	l_telcelular    = ""
	
	do until l_rs.eof

	  if CInt(l_rs("teldefault")) = -1 then
	     l_telparticular = l_rs("telnro")
	  end if

	  if CInt(l_rs("telcelular")) = -1 then
	     l_telcelular = l_rs("telnro")
	  end if

	  l_rs.movenext
	loop

	l_rs.Close

	'--- linea 1 ---
	if l_activo then
		linea =  "<td nowrap colspan='1'><b>Estado:&nbsp;</b>Activo</td>"
	else
		linea =  "<td nowrap colspan='1'><b>Estado:&nbsp;</b>Inactivo</td>"
	end if
	linea = linea & "<td nowrap colspan='1'><b>&nbsp;Telef. Interno:&nbsp;</b>" & l_empinterno & "</td>"
	linea = linea & "<td nowrap colspan='2'><b>&nbsp;e-mail:&nbsp;</b>" & l_empemail & "</td>"
	
	call imprimirLinea(linea, "basico")	

	'--- linea 2 ---
	linea =         "<td nowrap colspan='1'><b>Fecha Nac.:&nbsp;</b>" & l_terfecnac &"</td>"
	if CInt(l_tersex) then
	   linea = linea & "<td nowrap colspan='1'><b>&nbsp;Sexo:&nbsp;</b>Masculino</td>"
	else
	   linea = linea & "<td nowrap colspan='1'><b>&nbsp;Sexo:&nbsp;</b>Femenino</td>"
	end if
	linea = linea & "<td nowrap colspan='1'><b>&nbsp;Tipo y Nro. Doc.:&nbsp;</b>" & l_tidsigla & "-" & l_nrodoc & "</td>"
	linea = linea & "<td nowrap colspan='1'><b>&nbsp;CUIL:&nbsp;</b>" & l_cuil & "</td>"
	
	call imprimirLinea(linea, "basico")
	
	'--- linea 3 ---
	linea =         "<td nowrap colspan='1'><b>CAT:&nbsp;</b>" & l_nrocat & "</td>"
	linea = linea & "<td nowrap colspan='1'><b>&nbsp;Nacionalidad:&nbsp;</b>" & l_nacionaldes & "</td>"
	linea = linea & "<td nowrap colspan='1'><b>&nbsp;Edad:&nbsp;</b>" & l_edad & "</td>"
	linea = linea & "<td nowrap colspan='1'><b>&nbsp;Estado Civil:&nbsp;</b>" & l_estcivdesabr & "</td>"
	
	call imprimirLinea(linea, "basico")	

	'--- linea 4 ---
	linea =         "<td nowrap colspan='1'><b>Pais Naci.:&nbsp;</b>" & l_paisdesc & "</td>"
	linea = linea & "<td nowrap colspan='1'><b>&nbsp;Fecha Fin Cont.:&nbsp;</b>" & l_fecfincont & "</td>"
	if not l_tersex then
	   linea = linea & "<td nowrap colspan='1'><b>&nbsp;Ape. Casada:&nbsp;</b>" & l_tercasape & "</td>"
	else
	   linea = linea & "<td nowrap colspan='1'><b>&nbsp;Ape. Casada:&nbsp;</b></td>"
	end if
	linea = linea & "<td nowrap colspan='1'><b>&nbsp;Telef. Particular:&nbsp;</b>" & l_telparticular & "</td>"	

	call imprimirLinea(linea, "basico")		

	'--- linea 5 ---
	linea =  "<td nowrap colspan='2'><b>Reporta a:&nbsp;</b>" & l_repempleg & "&nbsp;-&nbsp;" & l_repterape & ", " & l_repternom & "</td>"
	linea =  linea & "<td nowrap colspan='2'><b>&nbsp;Telef. Celular:&nbsp;</b>" & l_telcelular & "</td>"	
	
	call imprimirLinea(linea, "basico")	
	
	call cerrar_tabla()
	
end sub 'datos_empleado


'-----------------------------------------------------------------------------------------------------------
'Imprime los datos de la organizacion
sub datos_organizacion()

    Dim l_tplatenro

	'Busco el modelo de estructuras default
	l_sql = " SELECT * FROM adptemplate WHERE tplatedefault = -1 "
	
	rsOpen l_rs, cn, l_sql, 0 
	
	l_tplatenro = ""
	
	if l_rs.eof then
	   if l_orden = "" then
	      l_orden = " ORDER BY htetdesde desc, estrdabr"
	   end if
	else
	   l_tplatenro = l_rs("tplatenro")
	end if
	
	l_rs.close
	
	if l_tplatenro = "" then
	
		l_sql = "select tipoestructura.tenro, tedabr, estrdabr, teorden, estructura.estrnro, claseestructura.cedabr "
		l_sql = l_sql & " , his_estructura.htetdesde, his_estructura.htethasta "
		l_sql = l_sql & " from his_estructura INNER JOIN tipoestructura ON his_estructura.tenro = tipoestructura.tenro "
		l_sql = l_sql & " INNER JOIN estructura ON estructura.estrnro = his_estructura.estrnro "
		l_sql = l_sql & " INNER JOIN claseestructura ON tipoestructura.cenro = claseestructura.cenro "
		l_sql = l_sql & " where his_estructura.ternro =" & l_ternro

		if trim(l_fechaestr)<> "" and not isnull(l_fechaestr) then
			l_sql = l_sql & " AND his_estructura.htetdesde <=" & cambiafecha(l_fechaestr,"","")
			l_sql = l_sql & " AND ((his_estructura.htethasta IS NULL) " 
			l_sql = l_sql & "      OR "
			l_sql = l_sql & "      (his_estructura.htethasta >=" & cambiafecha(l_fechaestr,"","")
			l_sql = l_sql & "       )) "
		end if

		l_sql = l_sql & " ORDER BY estrdabr"

	else
	
	    'Busco las que estan en el modelo
		l_sql = "SELECT tipoestructura.tenro, tedabr, estrdabr, teorden, estructura.estrnro, claseestructura.cedabr, tplaestrorden orden "
		l_sql = l_sql & " , his_estructura.htetdesde, his_estructura.htethasta "
		l_sql = l_sql & " FROM his_estructura INNER JOIN tipoestructura ON his_estructura.tenro = tipoestructura.tenro "
		l_sql = l_sql & " INNER JOIN estructura ON estructura.estrnro = his_estructura.estrnro "
		l_sql = l_sql & " INNER JOIN claseestructura ON tipoestructura.cenro = claseestructura.cenro "
		l_sql = l_sql & " LEFT JOIN adptte_estr ON tipoestructura.tenro = adptte_estr.tenro AND tplatenro = " & l_tplatenro
		l_sql = l_sql & " WHERE his_estructura.ternro =" & l_ternro

		if trim(l_fechaestr)<> "" and not isnull(l_fechaestr) then
			l_sql = l_sql & " AND his_estructura.htetdesde <=" & cambiafecha(l_fechaestr,"","")
			l_sql = l_sql & " AND ((his_estructura.htethasta IS NULL) " 
			l_sql = l_sql & "      OR "
			l_sql = l_sql & "      (his_estructura.htethasta >=" & cambiafecha(l_fechaestr,"","")
			l_sql = l_sql & "       )) "
		end if

		l_sql = l_sql & " AND NOT tplaestrorden IS NULL "
		
		l_sql = l_sql & " UNION "
		
	    'Busco las que no estan en el modelo	
		l_sql = l_sql & " SELECT tipoestructura.tenro, tedabr, estrdabr, teorden, estructura.estrnro, claseestructura.cedabr, 10000 orden "
		l_sql = l_sql & " , his_estructura.htetdesde, his_estructura.htethasta "
		l_sql = l_sql & " FROM his_estructura INNER JOIN tipoestructura ON his_estructura.tenro = tipoestructura.tenro "
		l_sql = l_sql & " INNER JOIN estructura ON estructura.estrnro = his_estructura.estrnro "
		l_sql = l_sql & " INNER JOIN claseestructura ON tipoestructura.cenro = claseestructura.cenro "
		l_sql = l_sql & " LEFT JOIN adptte_estr ON tipoestructura.tenro = adptte_estr.tenro AND tplatenro = " & l_tplatenro
		l_sql = l_sql & " WHERE his_estructura.ternro =" & l_ternro
	
		if trim(l_fechaestr)<> "" and not isnull(l_fechaestr) then
			l_sql = l_sql & " AND his_estructura.htetdesde <=" & cambiafecha(l_fechaestr,"","")
			l_sql = l_sql & " AND ((his_estructura.htethasta IS NULL) " 
			l_sql = l_sql & "      OR "
			l_sql = l_sql & "      (his_estructura.htethasta >=" & cambiafecha(l_fechaestr,"","")
			l_sql = l_sql & "       )) "
		end if

		l_sql = l_sql & " AND tplaestrorden IS NULL "
	
		l_sql = l_sql & " ORDER BY orden, tedabr, htetdesde DESC "
	
	end if
	
	rsOpen l_rs, cn, l_sql, 0 
	
	call titulo_organizacion()
	
	if l_rs.eof then
 	   call imprimirLinea("<td colspan='5'>No se encontraron datos.</td>", "orga")
	end if

	do until l_rs.eof
	    linea =         " <td nowrap>" & l_rs("tedabr") & "</td>"
	    linea = linea & " <td nowrap>" & l_rs("estrdabr") & "</td>"
	    linea = linea & " <td nowrap>" & l_rs("cedabr") & "</td>"
	    linea = linea & " <td nowrap align='center'>" & l_rs("htetdesde") & "</td>"
	    linea = linea & " <td nowrap align='center'>" & l_rs("htethasta") & "</td>"

        call imprimirLinea(linea, "orga")
		
		l_rs.MoveNext
	loop
	l_rs.Close
	
	call cerrar_tabla()

end sub 'datos_organizacion

'-----------------------------------------------------------------------------------------------------------
'Imprime los datos de las fases
sub datos_fases()

	l_sql = "SELECT altfec, bajfec, caudes, estado, empatareas "
	l_sql = l_sql & " FROM fases "
	l_sql = l_sql & " LEFT JOIN empant ON fases.empantnro=empant.empantnro "
	l_sql = l_sql & " LEFT JOIN causa ON fases.caunro=causa.caunro "
	l_sql = l_sql & " WHERE fases.empleado=" & l_ternro & " order by altfec DESC"

	rsOpen l_rs, cn, l_sql, 0 
	
	call titulo_fases()
	
	if l_rs.eof then
 	   call imprimirLinea("<td colspan='5'>No se encontraron datos.</td>", "fases")
	end if

	do until l_rs.eof
	    linea =         "<td align='center'>" & l_rs("altfec") & "</td>"
	    linea = linea & "<td align='center'>" & l_rs("bajfec") & "</td>"
	    linea = linea & "<td >" & l_rs("caudes") & "</td>"
		if CInt(l_rs("estado")) = 0 then
	       linea = linea & "<td align='center'>Inactivo</td>"
		else
	       linea = linea & "<td align='center'>Activo</td>"		
		end if
	    linea = linea & "<td align='center'>" & l_rs("empatareas") & "</td>"

        call imprimirLinea(linea, "fases")
		
		l_rs.MoveNext
	loop
	l_rs.Close
	
	call cerrar_tabla()

end sub 'datos_fases

'-----------------------------------------------------------------------------------------------------------
'Imprime los datos de los documentos
sub datos_documentos()

	l_sql = "SELECT tidnom, tidsigla, nrodoc, fecvtodoc "
	l_sql = l_sql & " FROM ter_doc "
	l_sql = l_sql & " INNER JOIN tipodocu ON tipodocu.tidnro = ter_doc.tidnro"	
	l_sql = l_sql & " WHERE ternro=" & l_ternro & " AND tipodocu.tidnro >= 5 AND tipodocu.tidnro NOT IN (10,20) ORDER BY tidnom DESC"

	rsOpen l_rs, cn, l_sql, 0 
	
	call titulo_documentos()
	
	if l_rs.eof then
 	   call imprimirLinea("<td colspan='4'>No se encontraron datos.</td>", "docu")
	end if

	do until l_rs.eof
	    linea =         "<td >" & l_rs("tidnom") & "</td>"
	    linea = linea & "<td >" & l_rs("tidsigla") & "</td>"
	    linea = linea & "<td >" & l_rs("nrodoc") & "</td>"
	    linea = linea & "<td >" & l_rs("fecvtodoc") & "</td>"

        call imprimirLinea(linea, "docu")
		
		l_rs.MoveNext
	loop
	l_rs.Close
	
	call cerrar_tabla()

end sub 'datos_documentos

'-----------------------------------------------------------------------------------------------------------
'Imprime los datos de los domicilios
sub datos_domicilios()

	l_sql = "SELECT tipodomi.tidodes,calle,nro,piso,oficdepto,torre,manzana,barrio,email "
	l_sql = l_sql & ",codigopostal,partnro,zonanro,cabdom.domdefault, detdom.domnro, locdesc "
	l_sql = l_sql & " FROM detdom INNER JOIN cabdom ON detdom.domnro=cabdom.domnro "
	l_sql = l_sql & " INNER JOIN tipodomi ON cabdom.tidonro=tipodomi.tidonro "
	l_sql = l_sql & " LEFT JOIN localidad on localidad.locnro = detdom.locnro"
	l_sql = l_sql & " WHERE cabdom.ternro=" & l_ternro
	if CInt(l_tipodom) > 0 then
	   l_sql = l_sql & " AND cabdom.tidonro=" & l_tipodom
	end if
	l_sql = l_sql & " ORDER BY tipodomi.tidodes DESC"

	rsOpen l_rs, cn, l_sql, 0 
	
	call titulo_domicilios()

	if l_rs.eof then
 	   call imprimirLinea("<td colspan='5'>No se encontraron datos.</td>", "domi")
	end if
	
	do until l_rs.eof
	    if CInt(l_rs("domdefault")) = -1 then
           linea = "<td nowrap align='center'>S&iacute;</td>"
		else
           linea = "<td nowrap align='center'>No</td>"
		end if
        
		linea = linea & "<td align='left'>" & l_rs("tidodes") & "</td>"
        linea = linea & "<td align='left'>" & l_rs("calle") & "&nbsp;-&nbsp;" & l_rs("nro") & "&nbsp;-&nbsp;" & l_rs("piso") & "</td>"
		linea = linea & "<td align='center'>" & l_rs("oficdepto") & "</td>"
        linea = linea & "<td nowrap align='left'>" & l_rs("locdesc") & "</td>"
        linea = linea & "<td align='right'>" & l_rs("codigopostal") & "</td>"
	
        call imprimirLinea(linea, "domi")
		
		l_rs.MoveNext
	loop
	l_rs.Close
	
	call cerrar_tabla()

end sub 'datos_domicilios

'-----------------------------------------------------------------------------------------------------------
'Imprime los datos de los familiares
sub datos_familiares()

	l_sql = "SELECT tercero.ternro,tercero.terape, tercero.ternom, famest, terfecnac, " &_
	      "famsalario, famfecvto, famCargaDGI, famDGIdesde, famDGIhasta, famemergencia, " &_ 
		  " paredesc, tidsigla, nrodoc " &_
	      "FROM  tercero INNER JOIN familiar ON tercero.ternro=familiar.ternro " &_
	      "LEFT JOIN ter_doc docu ON docu.ternro= tercero.ternro and docu.tidnro>0 and docu.tidnro<5 " &_
	      "LEFT JOIN tipodocu     ON tipodocu.tidnro= docu.tidnro " &_
		  "LEFT JOIN parentesco ON familiar.parenro=parentesco.parenro " &_
		  "WHERE familiar.empleado = " & l_ternro
		  
	if CStr(l_parentesco) <> "0" then
	   l_sql = l_sql & " AND familiar.parenro IN (" & l_parentesco & ") "
	end if
	l_sql = l_sql & " ORDER BY tercero.terape DESC"

	rsOpen l_rs, cn, l_sql, 0 
	
	call titulo_familiares()

	Dim l_famest
	Dim l_famsalario
	Dim l_famdgi
	Dim l_famemerg
	Dim l_famestudia
	Dim l_documento
	
	if l_rs.eof then
 	   call imprimirLinea("<td colspan='10'>No se encontraron datos.</td>", "fami")
	end if

	do until l_rs.eof
	    l_sql = "SELECT ternro "
   	    l_sql = l_sql & "FROM estudio_actual "
	    l_sql = l_sql & "WHERE ternro = " & l_rs("ternro")
		
		rsOpen l_rs2, cn, l_sql, 0
        if not l_rs2.eof then 
		    l_famestudia = "si"
		else
			l_famestudia = "no" 
        end if
		l_rs2.close
		
		l_documento = l_rs("tidsigla") & "-" &l_rs("nrodoc")
	
         if l_rs("famest") then l_famest = "Activo" else l_famest = "Inactivo" end if
         if l_rs("famsalario") then l_famsalario = "Si" else l_famsalario = "No" end if
		 if l_rs("famCargaDGI") then l_famdgi = "Si" else l_famdgi = "No" end if
		 if l_rs("famemergencia") then l_famemerg = "Si" else l_famemerg = "No" end if

        linea =         "<td nowrap align='left'>" & l_rs("paredesc") & "</td>"
        linea = linea & "<td nowrap align='left'>" & l_rs("terape") & ", " & l_rs("ternom") & "</td>"
		linea = linea & "<td nowrap align='center'>" & l_documento & "</td>"
		linea = linea & "<td nowrap align='center'>" & l_rs("terfecnac") & "</td>"
		linea = linea & "<td nowrap align='center'>" & edad(l_rs("terfecnac")) & "</td>"
		linea = linea & "<td nowrap align='center'>" & l_famestudia & "</td>"
        linea = linea & "<td nowrap align='left'>" & l_famest & "</td>"
        linea = linea & "<td nowrap align='center'>" &  l_famsalario & "</td>"
		linea = linea & "<td nowrap align='center'>" &  l_famDGI & "</td>"
		linea = linea & "<td nowrap align='right'>" & l_rs("famDGIdesde") & "-" & l_rs("famDGIhasta") & "</td>"
	
        call imprimirLinea(linea, "fami")
		
		l_rs.MoveNext
	loop
	l_rs.Close
	
	call cerrar_tabla()

end sub 'datos_familiares


'-----------------------------------------------------------------------------------------------------------
'Se encarga de imprimir una linea y si supera el largo de pagina empezar en una pagina nueva
sub imprimirLinea(strLinea, titulo)

	if l_nrolinea > l_Max_Lineas_X_Pag then 

	    if l_corte then
		   call cerrar_tabla()
           response.write "<tr><td nowrap style='page-break-before:always;background:white;' colspan=""" & l_cantColumnas & """><br></td></tr>"
		end if

		l_corte = true
		
	    call encabezado
		
		select case titulo
		   case "basico"
		      call titulo_empleado()
		   case "orga"
		      call titulo_organizacion()
		   case "fases"
		      call titulo_fases()
		   case "docu"
		      call titulo_documentos()
		   case "domi"
		      call titulo_domicilios()
		   case "fami"
		      call titulo_familiares()
		end select
		
		l_nropagina	= l_nropagina + 1
		l_nrolinea  = 1
	end if
	  
	response.write "<tr>"
	response.write strLinea
	response.write "</tr>"

	l_nrolinea  = l_nrolinea  + 1

end sub 'imprimirLinea(strLinea)

%>
<!DOCTYPE HTML PUBLIC "-//IETF//DTD HTML//EN">
<html>

<head>
<link href="../<%= c_estiloTabla %>" rel="StyleSheet" type="text/css">
<meta http-equiv="Content-Type" content="text/html; charset=iso-8859-1">
<title>Datos Empleado</title>
</head>
<script>
</script>

<body leftmargin="0" rightmargin="0" topmargin="0" bottommargin="0">
<%
Set l_rs = Server.CreateObject("ADODB.RecordSet")
Set l_rs2 = Server.CreateObject("ADODB.RecordSet")

l_nrolinea = 500
l_nropagina = 1
l_corte = false
l_cantColumnas = 15

'Busco los datos del empleado
l_sql = "SELECT * FROM empleado WHERE ternro=" & l_ternro

rsOpen l_rs, cn, l_sql, 0 

if not l_rs.eof then
   l_empnombre = l_rs("terape") & " " & l_rs("terape2") & ", " & l_rs("ternom") & " " & l_rs("ternom2")
   l_empleg    = l_rs("empleg")
end if

l_rs.close

response.write "<table cellpadding=""1"" cellspacing=""1"" style='border: 0px solid black;'>"  

'Muestro los datos basicos del empleado
call datos_empleado()

if CInt(l_fases) <> 0 then
   call datos_fases()
end if

if CInt(l_organizacion) <> 0 then
   call datos_organizacion()
end if

if CInt(l_documentos) <> 0 then
   call datos_documentos()
end if

if CInt(l_domicilios) <> 0 then
   call datos_domicilios()
end if

if CInt(l_familiares) <> 0 then
   call datos_familiares()
end if

set l_rs = Nothing
set l_rs2 = Nothing
cn.Close
set cn = Nothing

%>
</table>

<script>
  parent.ifrmListo=1;
</script>

</body>
</html>


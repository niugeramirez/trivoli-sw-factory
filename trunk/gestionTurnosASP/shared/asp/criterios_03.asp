<% Option Explicit %>
<!--#include virtual="/serviciolocal/shared/inc/sec.inc"-->
<!--#include virtual="/serviciolocal/shared/inc/const.inc"-->
<!--#include virtual="/serviciolocal/shared/db/conn_db.inc"-->
<!--#include virtual="/serviciolocal/shared/inc/fecha.inc"-->
<!--#include virtual="/serviciolocal/shared/inc/sqls.inc"-->
<!--#include virtual="/serviciolocal/shared/inc/util.inc"-->
<!--#include virtual="/serviciolocal/shared/inc/adovbs.inc"-->
<!--
-----------------------------------------------------------------------------
Archivo        : criterios_03.asp
Descripcion    : Modulo que se encarga del ABM de criterios - alta/modif
Creador        : Scarpa D.
Fecha Creacion : 28/11/2003
Modificacion   :16/09/2004 - Lisandro Moro - Correccion en la baja del sel_ter y recarga de ifrmfiltro
-----------------------------------------------------------------------------
-->
<% 
Dim l_tipo
Dim l_cm
Dim l_rs
Dim l_sql

Dim l_selnro
Dim l_selclase
Dim l_selsist
Dim l_selglobal
Dim l_seldesabr
Dim l_seldesext
Dim l_selsql
Dim l_selasp
Dim l_seltipnro

Dim l_seleccion
Dim l_arr
Dim l_arr2
Dim l_i
Dim l_ternro


l_tipo = request.QueryString("tipo")

l_selnro    = request("selnro")
l_selclase  = request("selclase")
l_selsist   = getCheckbox("selsist")
l_selglobal = getCheckbox("selglobal")
l_seldesabr = request("seldesabr")
l_seldesext = request("seldesext")
l_selsql    = request("sql")
l_selasp    = request("asp")
l_seltipnro = request("seltipnro")
l_seleccion = request("seleccion")

if l_selclase = 1 then
   l_selasp = ""
   l_seleccion = ""
else
  if l_selclase = 2 then
     l_selsql = ""
     l_seleccion = ""	 
  else
     l_selsql = ""
     l_selasp = ""
  end if
end if

%>
<script src="/serviciolocal/shared/js/fn_fechas.js"></script>
<%

set l_cm = Server.CreateObject("ADODB.Command")
Set l_rs = Server.CreateObject("ADODB.RecordSet")

cn.beginTrans

if l_selnro <> "" then
 
	'Borro todos los datos de sel_ter
	l_sql = "SELECT * FROM sel_ter WHERE selnro =" & l_selnro
	
	'rsOpen l_rs, cn, l_sql, 0 
	rsOpenCursor l_rs, cn, l_sql,0,adOpenKeyset
	l_cm.activeconnection = Cn
	do until l_rs.eof
	
		l_sql = "DELETE FROM sel_ter WHERE ternro=" & l_rs("ternro") & " AND selnro = " & l_selnro
		
		l_cm.CommandText = l_sql
		cmExecute l_cm, l_sql, 0
	
	    l_rs.MoveNext
	loop
	l_rs.Close

end if	
	
'Guardo los datos
if (l_tipo = "A") OR (l_tipo = "TA") OR (l_tipo = "LE") then 
	l_sql = "INSERT INTO seleccion "
	l_sql = l_sql & "(seldesabr, seldesext, seltipnro, selclase , selprog , selsql , selsist, selglobal) "
	l_sql = l_sql & " values ('" & l_seldesabr & "','" & l_seldesext & "', " 
	l_sql = l_sql & l_seltipnro & "," 
	l_sql = l_sql & l_selclase & ",'" & l_selasp & "','" & l_selsql & "'," & l_selsist & "," & l_selglobal & ")"
	
	l_cm.activeconnection = Cn
	l_cm.CommandText = l_sql
	cmExecute l_cm, l_sql, 0

	'Busco el codigo del recien insertado
	l_sql = fsql_seqvalue("next_id","seleccion")
	rsOpen l_rs, cn, l_sql, 0
	l_selnro =l_rs("next_id")
	l_rs.close

    'Me fijo si hay que guardar la lista de empleados
	if CInt(l_selclase) = 3 then
  	  l_arr = Split(l_seleccion,",")		
   	  l_i = 1
	
	  do while l_i <= UBound(l_arr)
	    if instr(1,l_arr(l_i),"@") then
		   l_arr2 = split(l_arr(l_i),"@")
		   l_ternro = l_arr2(0)
		else
		   l_ternro = l_arr(l_i)
		end if

        'Inserto el dato en la BD
	    l_sql = "INSERT INTO sel_ter "
	    l_sql = l_sql & "(ternro, selnro)"
	    l_sql = l_sql & " values (" & l_ternro  & ", " & l_selnro & ")"
	
	    l_cm.CommandText = l_sql
	    cmExecute l_cm, l_sql, 0

 		l_i = l_i + 1
	  loop
	
	end if
	
else
	l_sql = "UPDATE seleccion SET "
	l_sql = l_sql & " seldesabr ='" & l_seldesabr & "',"
	l_sql = l_sql & " seldesext ='" & l_seldesext & "', "
	l_sql = l_sql & " seltipnro = " & l_seltipnro & ","
	l_sql = l_sql & " selclase = "  & l_selclase  & ","
	l_sql = l_sql & " selprog = '"  & l_selasp    & "',"
	l_sql = l_sql & " selsql = '"   & l_selsql    & "',"
	l_sql = l_sql & " selsist = "   & l_selsist   & ","
	l_sql = l_sql & " selglobal = " & l_selglobal & ""
	l_sql = l_sql & " WHERE selnro = " & l_selnro
	
	l_cm.activeconnection = Cn
	l_cm.CommandText = l_sql
	cmExecute l_cm, l_sql, 0

    'Me fijo si hay que guardar la lista de empleados
	if CInt(l_selclase) = 3 then
  	  l_arr = Split(l_seleccion,",")		
   	  l_i = 1
	
	  do while l_i <= UBound(l_arr)
	    if instr(1,l_arr(l_i),"@") then
		   l_arr2 = split(l_arr(l_i),"@")
		   l_ternro = l_arr2(0)
		else
		   l_ternro = l_arr(l_i)
		end if

        'Inserto el dato en la BD
	    l_sql = "INSERT INTO sel_ter "
	    l_sql = l_sql & "(ternro, selnro)"
	    l_sql = l_sql & " values (" & l_ternro  & ", " & l_selnro & ")"
	
	    l_cm.CommandText = l_sql
	    cmExecute l_cm, l_sql, 0

 		l_i = l_i + 1
	  loop
	
	end if
	
end if

cn.commitTrans

Set cn = Nothing
Set l_cm = Nothing
%>

<script>
  alert('Operación Realizada.');
  if (window.opener.opener.ifrmfiltros()){
  	window.opener.opener.ifrmfiltros.location.reload();
  }
  window.opener.close();
  window.close();
</script>

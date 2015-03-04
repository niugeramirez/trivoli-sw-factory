<!--#include virtual="/turnos/shared/inc/users_sql.inc"-->
<%
function usuario(tipo, id, clave)
'	response.write "<script>alert('" & tipo & " - " & id & " - " & clave & "')</script>"
	select case tipo
		case 1
			'Alta de usuario
			call crear_usuario(id, clave)
		case 2
			'Baja de usuario
			call borrar_usuario(id)
		case 3
			'Modificacion de usuario
			call cambiar_password(id, clave)
   end select
end function

function fsql_first(sql,cant)
  Dim auxi
  Select Case l_base
        Case "1" 
			auxi = right(sql,len(sql) - (Instr(ucase(sql), "SELECT") + 6))
			auxi = "SELECT" & " FIRST " & cant & " " & auxi
        Case "2" 
			auxi = right(sql,len(sql) - (Instr(ucase(sql), "SELECT") + 6))
			auxi = "SELECT TOP " & cant & " " & auxi
        Case "3" 
			auxi = sql & " fetch first " & cant & " rows only "
        Case "8" 
			auxi = right(sql,len(sql) - (Instr(ucase(sql), "SELECT") + 6))
			auxi = "SELECT TOP " & cant & " " & auxi
	Case "9" 
			auxi = right(sql,len(sql) - (Instr(ucase(sql), "SELECT") + 6))
			auxi = "SELECT TOP " & cant & " " & auxi			
        Case "10" 
			auxi = right(sql,len(sql) - (Instr(ucase(sql), "SELECT") + 6))
			auxi = "SELECT TOP " & cant & " " & auxi
	Case "12" 
			auxi = right(sql,len(sql) - (Instr(ucase(sql), "SELECT") + 6))
			auxi = "SELECT TOP " & cant & " " & auxi
	Case "13" 
			auxi = right(sql,len(sql) - (Instr(ucase(sql), "SELECT") + 6))
			auxi = "SELECT TOP " & cant & " " & auxi
	Case "14" 
			auxi = right(sql,len(sql) - (Instr(ucase(sql), "SELECT") + 6))
			auxi = "SELECT TOP " & cant & " " & auxi
    Case "15" 
			auxi = right(sql,len(sql) - (Instr(ucase(sql), "SELECT") + 6))
			auxi = "SELECT TOP " & cant & " " & auxi

  End Select
  fsql_first = auxi
end function

function fsql_seqvalue2(ByVal campo, ByVal tabla)
  Dim auxi
  Select Case l_base
        Case "1"
			auxi = "select unique DBINFO('sqlca.sqlerrd1') as " & campo & " from " & tabla & " "
        Case "2" 
			auxi = "select @@IDENTITY as " & campo & " "
        Case "3" 
			auxi = "SELECT IDENTITY_VAL_LOCAL() as " & campo & " FROM SYSIBM.SYSDUMMY1 "
        Case "8" 
			auxi = "select @@IDENTITY as " & campo & " "
        Case "9" 
			auxi = "select @@IDENTITY as " & campo & " "			
        Case "10" 
			auxi = "select @@IDENTITY as " & campo & " "
        Case "12" 
			auxi = "select @@IDENTITY as " & campo & " "
        Case "13" 
			auxi = "select @@IDENTITY as " & campo & " "
        Case "14" 
			auxi = "select @@IDENTITY as " & campo & " "
        Case "15" 
			auxi = "select @@IDENTITY as " & campo & " "

  End Select
  fsql_seqvalue = auxi
end function

function fsql_seqvalue(ByVal campo, ByVal tabla)
  Dim auxi
  Select Case l_base
        Case "1"
			auxi = "select unique DBINFO('sqlca.sqlerrd1') as " & campo & " from " & tabla & " "
        Case "2" 
			auxi = "select IDENT_CURRENT ('" & tabla & "') as " & campo & " "
        Case "3" 
			auxi = "SELECT IDENTITY_VAL_LOCAL() as " & campo & " FROM SYSIBM.SYSDUMMY1 "
        Case "8" 
			auxi = "select @@IDENTITY as " & campo & " "
        Case "9" 
			auxi = "select @@IDENTITY as " & campo & " "			
        Case "10" 
			auxi = "select @@IDENTITY as " & campo & " "
        Case "12" 
			auxi = "select @@IDENTITY as " & campo & " "
        Case "13" 
			auxi = "select @@IDENTITY as " & campo & " "
        Case "14" 
			auxi = "select @@IDENTITY as " & campo & " "
        Case "15" 
			auxi = "select @@IDENTITY as " & campo & " "

  End Select
  fsql_seqvalue2 = auxi
end function

function fsql_query(query)
Dim Sql
Select Case l_base
	Case "1" 
 		Select Case query
		Case "0001"
			sql = "select tipoestructura.tenro, tedabr, estrdabr, teorden, estructura.estrnro "
			sql = sql & "from tipoestructura, outer(estruc_actual, estructura) "
			sql = sql & "where ternro = " & ternro & " "
			sql = sql & "and estruc_actual.tenro = tipoestructura.tenro "
			sql = sql & "and tipoestructura.cenro = " & cenro & " "
			sql = sql & "and estructura.estrnro = estruc_actual.estrnro "
			sql = sql & "order by teorden"
		End Select
	Case "2" 
 		Select Case query
		Case "0001"
			sql = "select tipoestructura.tenro, tedabr, estrdabr, teorden, estructura.estrnro "
			sql = sql & "FROM estruc_actual "
			sql = sql & "INNER JOIN estructura ON estruc_actual.estrnro = estructura.estrnro "
			sql = sql & "RIGHT OUTER JOIN tipoestructura ON estruc_actual.tenro = tipoestructura.tenro AND estruc_actual.ternro = " & ternro & " "
			sql = sql & "WHERE tipoestructura.cenro = " & cenro & " "
			sql = sql & "order by teorden"
		End Select
    Case "3" 
 		Select Case query
		Case "0001"
			sql = "select tipoestructura.tenro, tedabr, estrdabr, teorden, estructura.estrnro "
			sql = sql & "FROM estruc_actual "
			sql = sql & "INNER JOIN estructura ON estruc_actual.estrnro = estructura.estrnro "
			sql = sql & "RIGHT OUTER JOIN tipoestructura ON estruc_actual.tenro = tipoestructura.tenro AND estruc_actual.ternro = " & ternro & " "
			sql = sql & "WHERE tipoestructura.cenro = " & cenro & " "
			sql = sql & "order by teorden"
		End Select
	Case "8" 
 		Select Case query
		Case "0001"
			sql = "select tipoestructura.tenro, tedabr, estrdabr, teorden, estructura.estrnro "
			sql = sql & "FROM estruc_actual "
			sql = sql & "INNER JOIN estructura ON estruc_actual.estrnro = estructura.estrnro "
			sql = sql & "RIGHT OUTER JOIN tipoestructura ON estruc_actual.tenro = tipoestructura.tenro AND estruc_actual.ternro = " & ternro & " "
			sql = sql & "WHERE tipoestructura.cenro = " & cenro & " "
			sql = sql & "order by teorden"
		End Select
	Case "9" 
 		Select Case query
		Case "0001"
			sql = "select tipoestructura.tenro, tedabr, estrdabr, teorden, estructura.estrnro "
			sql = sql & "FROM estruc_actual "
			sql = sql & "INNER JOIN estructura ON estruc_actual.estrnro = estructura.estrnro "
			sql = sql & "RIGHT OUTER JOIN tipoestructura ON estruc_actual.tenro = tipoestructura.tenro AND estruc_actual.ternro = " & ternro & " "
			sql = sql & "WHERE tipoestructura.cenro = " & cenro & " "
			sql = sql & "order by teorden"
		End Select		
	Case "10" 
 		Select Case query
		Case "0001"
			sql = "select tipoestructura.tenro, tedabr, estrdabr, teorden, estructura.estrnro "
			sql = sql & "FROM estruc_actual "
			sql = sql & "INNER JOIN estructura ON estruc_actual.estrnro = estructura.estrnro "
			sql = sql & "RIGHT OUTER JOIN tipoestructura ON estruc_actual.tenro = tipoestructura.tenro AND estruc_actual.ternro = " & ternro & " "
			sql = sql & "WHERE tipoestructura.cenro = " & cenro & " "
			sql = sql & "order by teorden"
		End Select
	Case "12" 
 		Select Case query
		Case "0001"
			sql = "select tipoestructura.tenro, tedabr, estrdabr, teorden, estructura.estrnro "
			sql = sql & "FROM estruc_actual "
			sql = sql & "INNER JOIN estructura ON estruc_actual.estrnro = estructura.estrnro "
			sql = sql & "RIGHT OUTER JOIN tipoestructura ON estruc_actual.tenro = tipoestructura.tenro AND estruc_actual.ternro = " & ternro & " "
			sql = sql & "WHERE tipoestructura.cenro = " & cenro & " "
			sql = sql & "order by teorden"
		End Select
	Case "13" 
 		Select Case query
		Case "0001"
			sql = "select tipoestructura.tenro, tedabr, estrdabr, teorden, estructura.estrnro "
			sql = sql & "FROM estruc_actual "
			sql = sql & "INNER JOIN estructura ON estruc_actual.estrnro = estructura.estrnro "
			sql = sql & "RIGHT OUTER JOIN tipoestructura ON estruc_actual.tenro = tipoestructura.tenro AND estruc_actual.ternro = " & ternro & " "
			sql = sql & "WHERE tipoestructura.cenro = " & cenro & " "
			sql = sql & "order by teorden"
		End Select
		
End Select
fsql_query = sql
end function
%>

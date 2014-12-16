<% Option Explicit %>
<!--#include virtual="/ticket/shared/db/conn_db.inc"-->
<%
'on error goto 0

dim Actual
dim perfil
dim username
dim btnaccess
dim myRs
dim sql
dim autorizado
dim l_menu

Set myRs = Server.CreateObject("ADODB.RecordSet")

username = UCase(Session("Username"))
l_menu	 = request.QueryString("menu")

sql = "SELECT perfnom FROM user_per"
sql = sql & " inner join perf_usr on perf_usr.perfnro = user_per.perfnro"
sql = sql & " where upper(iduser) = '" & username & "'"
rsOpen myRs, cn, sql, 0
perfil = myRs("perfnom")
myRs.Close

sql = "SELECT menuname, menuaccess, action from menumstr "
sql = sql & " WHERE  parent = 'Rhpro'"

rsOpen myRs, cn, sql, 0
myRs.MoveFirst

Dim arreglo
Dim i

function menu
	if LCase(myRs("action")) <> "" then
		%><script>
		if (parent.document.all.<%= LCase(myRs("menuname")) %>){
			parent.document.all.<%= LCase(myRs("menuname")) %>.style.filter = 'none';
			parent.document.all.a<%= LCase(myRs("menuname")) %>.href = "Javascript:parent.ventanas('../<%= LCase(myRs("action")) %>','<%= myRs("menuname") %>');";
			}
		</script><%
	end if
end function

do until myRs.eof
	if not username = "SUPER" then
         if ( myRs("menuaccess") = perfil )  or ( myRs("menuaccess") = "*" ) then
	        'imprimo las rutas de los modulos
		    response.write "&" & LCase(myRs("menuname")) & "=" & LCase(myRs("action"))   & "&"
			if l_menu = "html" then
				menu
			end if
		 else
		    arreglo = split(myRs("menuaccess"),";")
			
			for i=0 to UBound(arreglo)
			   if UCase(arreglo(i)) = UCase(perfil) then
			      'imprimo las rutas de los modulos
			      response.write "&" & LCase(myRs("menuname")) & "=" & LCase(myRs("action"))
		  			if l_menu = "html" then
						menu
					end if
			   end if
			next
	 
	    end if

	else
   	       'imprimo las rutas de los modulos
			response.write "&" & LCase(myRs("menuname")) & "=" & LCase(myRs("action")) & "&"
			if l_menu = "html" then
				menu
			end if
	end if
	myRs.MoveNext
loop

response.write "&"

myRs.Close
set myRs = Nothing
%>

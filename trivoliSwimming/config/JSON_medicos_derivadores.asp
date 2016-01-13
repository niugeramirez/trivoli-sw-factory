<% Option Explicit %>
<!--#include virtual="/turnos/shared/inc/sec.inc"-->
<!--#include virtual="/turnos/shared/inc/const.inc"-->
<!--#include virtual="/turnos/shared/db/conn_db.inc"-->
<% 


Dim keywords
Dim keywords_cmd
Dim output 






Set keywords_cmd = Server.CreateObject ("ADODB.Command")
keywords_cmd.ActiveConnection = cn
keywords_cmd.CommandText = "SELECT id, nombre as name "
keywords_cmd.CommandText = keywords_cmd.CommandText & " FROM medicos_derivadores"
keywords_cmd.CommandText = keywords_cmd.CommandText & " where nombre like '%"  & Request.QueryString("term") & "%'"
keywords_cmd.CommandText = keywords_cmd.CommandText &  " and empnro = " & Session("empnro")
keywords_cmd.Prepared = true
Set keywords = keywords_cmd.Execute

output = "["

While (NOT keywords.EOF) 
    output = output & "{""id"":""" & keywords.Fields.item("id") & """,""value"":""" & keywords.Fields.Item("name") & """},"
     keywords.MoveNext()
Wend

keywords.Close()
Set keywords = Nothing

output=Left(output,Len(output)-1)
output = output & "]"
response.write output


%>





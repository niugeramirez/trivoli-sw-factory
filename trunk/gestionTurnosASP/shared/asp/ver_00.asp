<% Option Explicit %>
<!--
-----------------------------------------------------------------------------
Archivo        : ver_00.asp
Descripcion    : Modulo que se encarga de loguearse y mostrar un url
Creador        : Scarpa D.
Fecha Creacion : 03/08/2004
Modificacion   :
-----------------------------------------------------------------------------
-->
<%
'Ejemplo de uso:
'http://127.0.0.1/serviciolocal/shared/asp/ver_00.asp?user=sa&pass=&base=8&url=http://127.0.0.1/serviciolocal/liq/rep_recibo_liq_03.asp?bpronro=1734

  if request("user") = "" OR request("base") = "" then
     response.end
  end if

  Session("password") = request("pass")
  Session("username") = request("user")
  Session("base") = request("base")
  Session("Time") = now
%>
<!--#include virtual="/serviciolocal/shared/db/conn_db.inc"-->
<%
on error goto 0

Dim l_url

l_url = request("url")

%>

<script>
  window.location = '<%= l_url%>';
</script>

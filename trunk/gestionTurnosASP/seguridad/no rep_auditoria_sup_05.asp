<!--#include virtual="/serviciolocal/shared/db/conn_db.inc"-->
<!--#include virtual="/serviciolocal/shared/inc/adovbs.inc"-->
<!-----------------------------------------------------------------------------
Archivo		: rep_auditoria_sup_05.asp
Autor		: JMH
Creacion	: 24/01/2005
Descripcion	: Modulo que se encarga de borrar un histórico
-------------------------------------------------------------------------------
-->
<% 
'on error goto 0
 
 Dim l_rs
 Dim l_cm
 Dim l_sql
 
 Dim l_bpronro
 Dim l_anchoselecthist
 
 l_bpronro			= Request.QueryString("bpronro")
 l_anchoselecthist  = Request.QueryString("anchoselecthist")
 
 cn.beginTrans
 
 Set l_rs = Server.CreateObject("ADODB.RecordSet") 
 Set l_cm = Server.CreateObject("ADODB.Command")
 
 l_sql =           " SELECT repnro "
 l_sql  = l_sql  & " FROM rep_auditoria "
 l_sql  = l_sql  & " WHERE bpronro = " & l_bpronro
 
 rsOpenCursor l_rs, cn, l_sql, 0, adOpenKeyset
 
 do until l_rs.eof
	l_sql =           " DELETE "
	l_sql  = l_sql  & " FROM rep_auditoria "
	l_sql  = l_sql  & " WHERE repnro = " & l_rs("repnro")
	
	l_cm.activeconnection = Cn
    l_cm.CommandText = l_sql
    cmExecute l_cm, l_sql, 0
	
    l_rs.Movenext
 loop
 
 l_rs.Close
 
 cn.commitTrans
 
%>	
<script>
	alert('Operación Realizada.');
	opener.ifrm.location = "blanc.html";
	opener.ifrmauditoria.location="combo_hist_auditoria_sup_00.asp?ancho=<%= l_anchoselecthist %>";
	window.close(); 
</script>

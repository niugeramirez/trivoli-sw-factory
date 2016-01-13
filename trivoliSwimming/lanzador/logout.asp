<% Session.Abandon 
if request("arg") = 1 then %>
  	   <script>window.close();</script>
<%	end if %>

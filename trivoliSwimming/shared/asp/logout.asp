<% Session.Abandon 
if request("arg") = 1 then %>
  	   <script>//window.opener = null;
	           //window.close();
			    window.open("","_parent","");
                window.close(); 
			   </script>
<%	end if %>

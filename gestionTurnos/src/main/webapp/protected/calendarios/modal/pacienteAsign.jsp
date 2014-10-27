<%@ page contentType="text/html;charset=UTF-8" language="java" %>
<%@ taglib prefix="spring" uri="http://www.springframework.org/tags" %>

<!-- DIV con el dialogo de busqueda -->
<div id="pacienteAsign"
     class="modal fade"
     role="dialog"
     aria-hidden="true">
     
   	<!-- 		DIV del  diálogo para el AJAX request -->
	<div id="loadingModalPacientes" class="modal hide fade in centering" role="dialog" aria-hidden="true">
           <div id="divLoadingIcon" class="text-center">
               <div class="glyphicon glyphicon-align-center loading"></div>
           </div>
    </div>
	                
	<jsp:include page="subPacienteSearch.jsp"/>
	<jsp:include page="subPacienteSearchResult.jsp"/>				
 </div>
<%@ page contentType="text/html;charset=UTF-8" language="java" %>
<%@ taglib prefix="spring" uri="http://www.springframework.org/tags" %>

<!-- DIV con el dialogo de eliminar -->
<div id="eliminarObrasSocialesDialog"
    class="modal fade"
    role="dialog"
    aria-hidden="true">
     
	<div class="modal-dialog">
   		<div class="modal-content">
    		<div class="modal-header">
    			<h4 id="eliminarObrasSocialesDialogLabel" class="displayInLine">
            		<spring:message code="delete"/>&nbsp;<spring:message code="obraSocial"/>
        		</h4>
   			</div>
    	
			<div class="modal-body">
				<p><spring:message code="delete.confirm"/>:&nbsp;{{obraSocial.nombre}}?</p>
				
				<!--Mensaje de error-->
				<span class="alert alert-danger"
		          ng-show="errorSubmit">
		        	<spring:message code="request.error"/>
		    	</span>
    			<span class="alert alert-danger"
          		ng-show="errorAccesoIlegal">
        			<spring:message code="request.illegal.access"/>
    			</span>
			</div>
			
			<!--     		DIV con los botones de eliminar y cancelar -->
			<div class="modal-footer">
			 	<form name="eliminarObraSocialForm" novalidate>
            	<input type="submit"
                   class="btn btn-primary"
                   ng-click="eliminarObraSocial();"
                   value='<spring:message code="delete"/>'/>
            	<button class="btn btn-default"
                    data-dismiss="modal"
                    ng-click="exit('#eliminarObrasSocialesDialog');"
                    aria-hidden="true">
                	<spring:message code="cancel"/>
            	</button>
       			 </form>
			</div>    	
		</div>
    </div>
</div>

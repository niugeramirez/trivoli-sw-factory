<%@ page contentType="text/html;charset=UTF-8" language="java" %>
<%@ taglib prefix="spring" uri="http://www.springframework.org/tags" %>

<!-- DIV con el dialogo de busqueda -->
<div id="buscarObrasSocialesDialog"
     class="modal fade"
     role="dialog"
     aria-hidden="true">
    
    <div class="modal-dialog  modal-lg">
   		<div class="modal-content"> 
			<div class="modal-header">
        		<h3 id="buscarObrasSocialesModalLabel" class="displayInLine">
            		<spring:message code="search"/>
        		</h3>
    		</div>
			
			<!--     	DIV contenedor de los parametros de busqueda y los mensajes de error correspondientes -->
    		<div class="modal-body">
        		<form name="buscarObrasSocialesForm" novalidate class="form-horizontal">
					
					<!--         			DIV con la parte de busqueda -->
        			<div class="form-group" ng-class="{'has-error': mostrarErrorValidacion && buscarObrasSocialesForm.filtroNombre.$error.required}">
        				<label for="txtNombre" class="col-lg-3 control-label"><spring:message code="search.for"/>:</label>
            			
	            		<div class="col-lg-7">
		                    	<input type="text"
		                    		class="form-control"
						    		id="txtNombre"	
		                           	autofocus
		                           	required
		                           	ng-model="filtroNombre"
		                           	name="filtroNombre"
		                           	placeholder="<spring:message code='obrasSociales.nombre'/> "/>
	                	</div>
	                	<div class="col-lg-2">
	                    	<label class="displayInLine">
	                        	<span class="alert alert-danger help-block"
	                              		ng-show="mostrarErrorValidacion && buscarObrasSocialesForm.filtroNombre.$error.required">
	                            		<spring:message code="required"/>
	                        	</span>
	                    	</label>
	                	</div>
	            	</div>
            	</form>
        		<span class="alert alert-danger"
		          ng-show="errorSubmit">
		        	<spring:message code="request.error"/>
			    </span>
	    		<span class="alert alert-danger"
	          		ng-show="errorAccesoIlegal">
	        		<spring:message code="request.illegal.access"/>
	    		</span>
    		</div>
    		
			<!--     		DIV con los botones de buscar y cancelar -->
    		<div class="modal-footer">
    			<input type="submit"
                   		class="btn btn-primary"
                   		ng-click="buscarObrasSociales(buscarObrasSocialesForm, false);"
                   		value='<spring:message code="search"/>'
                    	/>
            		<button class="btn btn-default"
                    	data-dismiss="modal"
                    	ng-click="exit('#buscarObrasSocialesDialog');"
                    	aria-hidden="true">
                		<spring:message code="cancel"/>
            		</button>
    		</div>
    	</div>
    </div>			
 </div>
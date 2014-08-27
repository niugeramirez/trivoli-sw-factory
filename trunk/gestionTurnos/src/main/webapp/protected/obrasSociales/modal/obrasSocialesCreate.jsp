<%@ page contentType="text/html;charset=UTF-8" language="java" %>
<%@ taglib prefix="spring" uri="http://www.springframework.org/tags" %>

<!-- DIV con el dialogo de Creacion -->

<div id="crearObrasSocialesDialog"
     class="modal fade"
     role="dialog"
     aria-hidden="true">
    
    <div class="modal-dialog modal-lg">
   		<div class="modal-content">
    
    		<div class="modal-header">
        		<h4 id="crearObraSocialDialogLabel" class="displayInLine">
            		<spring:message code="create"/>&nbsp;<spring:message code="obraSocial"/>
        		</h4>
    		</div>
    
    		<div class="modal-body">
        		<form name="crearObraSocialForm" novalidate class="well form-horizontal">             		
                	
                	<div class="form-group" ng-class="{'has-error': mostrarErrorValidacion && crearObraSocialForm.nombre.$error.required}">
						<label for="txtDescripcion" class="col-lg-3 control-label"><spring:message code="obrasSociales.nombre"/>:</label>
						<div class="col-lg-7">
							<input type="text"
						    	class="form-control"
						    	id="txtDescripcion"
                               	required
                               	autofocus
                               	ng-model="obraSocial.nombre"
                               	name="nombre"
                               	placeholder="<spring:message code='obrasSociales.nombre'/>"/>         
						</div>
						<div class="col-lg-2">
							<span class="alert alert-danger help-block"
                                      ng-show="mostrarErrorValidacion && crearObraSocialForm.nombre.$error.required">
                                        <spring:message code="required"/>
                                </span>
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
    		
    		<!--     		DIV con los botones de crear y cancelar -->
    		<div class="modal-footer">
    			 <input type="submit"
                       		class="btn btn-primary"
                       		ng-click="crearObraSocial(crearObraSocialForm);"
                       		value='<spring:message code="create"/>'/>
                	<button class="btn btn-default"
                        	data-dismiss="modal"
                        	ng-click="exit('#crearObrasSocialesDialog');"
                        	aria-hidden="true">
                    		<spring:message code="cancel"/>
                	</button>
    		</div>
		</div>
	</div>
</div>

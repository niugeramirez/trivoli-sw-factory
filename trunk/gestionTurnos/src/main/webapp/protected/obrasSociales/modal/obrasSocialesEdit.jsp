<%@ page contentType="text/html;charset=UTF-8" language="java" %>
<%@ taglib prefix="spring" uri="http://www.springframework.org/tags" %>

<!-- DIV con el dialogo de Editar -->

<div id="editarObrasSocialesDialog"
     class="modal fade"
     role="dialog"
     aria-hidden="true">
	 <div class="modal-dialog modal-lg">
   		<div class="modal-content">     
			<!-- 		    DIV con el titulo(cebecera) del modal -->
		    <div class="modal-header">
		        <h3 id="editarObraSocialModalLabel" class="displayInLine">
		            <spring:message code="update"/>&nbsp;<spring:message code="obraSocial"/>
		        </h3>
		    </div>
		    
			<!-- 		    DIV con el cuerpo del modal -->
		    <div class="modal-body">
		        <form name="editarObraSocialForm" novalidate class="well form-horizontal">
<!-- 		            <input type="hidden" -->
<!-- 		                   required -->
<!-- 		                   ng-model="recurso.id" -->
<!-- 		                   name="id" -->
<!-- 		                   value="{{recurso.id}}"/> -->
					
					<div class="form-group" ng-class="{'has-error': mostrarErrorValidacion && editarObraSocialForm.nombre.$error.required}">
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
							<label>
                            	<span class="alert alert-danger help-block"
                                      ng-show="mostrarErrorValidacion && editarObraSocialForm.nombre.$error.required">
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
			
			<!-- 			Div con el footer del modal -->
			<div class="modal-footer">
				 <input type="submit"
			     	class="btn btn-primary"
			        ng-click="updateObraSocial(editarObraSocialForm);"
			        value='<spring:message code="update"/>'/>
			     <button class="btn btn-default"
			     	data-dismiss="modal"
			        ng-click="exit('#editarObrasSocialesDialog');"
			        aria-hidden="true">
			        <spring:message code="cancel"/>
			     </button>
			</div>
		</div>
	</div>				    
</div>

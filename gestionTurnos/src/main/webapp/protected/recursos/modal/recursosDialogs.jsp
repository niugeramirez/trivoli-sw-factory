<%@ page contentType="text/html;charset=UTF-8" language="java" %>
<%@ taglib prefix="spring" uri="http://www.springframework.org/tags" %>


<div id="crearRecursoDialog"
     class="modal fade"
     role="dialog"
     aria-hidden="true">
    
    <div class="modal-dialog modal-lg">
   		<div class="modal-content">
    
    		<div class="modal-header">
        		<h4 id="crearRecursoDialogLabel" class="displayInLine">
            		<spring:message code="create"/>&nbsp;<spring:message code="recurso"/>
        		</h4>
    		</div>
    
    		<div class="modal-body">
        		<form name="crearRecursoForm" novalidate class="well form-horizontal">             		
                	
                	<!--         			DIV con la parte de busqueda -->
                	<div class="form-group" ng-class="{'has-error': mostrarErrorValidacion && crearRecursoForm.descripcion.$error.required}">
						<label for="txtDescripcion" class="col-lg-3 control-label"><spring:message code="recursos.descripcion"/>:</label>
						<div class="col-lg-7">
							<input type="text"
						    	class="form-control"
						    	id="txtDescripcion"
                               	required
                               	autofocus
                               	ng-model="recurso.descripcion"
                               	name="descripcion"
                               	placeholder="<spring:message code='recursos.descripcion'/>"/>         
						</div>
						<div class="col-lg-2">
							<span class="alert alert-danger help-block"
                                      ng-show="mostrarErrorValidacion && crearRecursoForm.descripcion.$error.required">
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
    		
    		<!--     		DIV con los botones de buscar y cancelar -->
    		<div class="modal-footer">
    			 <input type="submit"
                       		class="btn btn-primary"
                       		ng-click="crearRecurso(crearRecursoForm);"
                       		value='<spring:message code="create"/>'/>
                	<button class="btn btn-default"
                        	data-dismiss="modal"
                        	ng-click="exit('#crearRecursoDialog');"
                        	aria-hidden="true">
                    		<spring:message code="cancel"/>
                	</button>
    		</div>
		</div>
	</div>
</div>

<div id="editarRecursoDialog"
     class="modal fade"
     role="dialog"
     aria-hidden="true">
	 <div class="modal-dialog modal-lg">
   		<div class="modal-content">     
		    <div class="modal-header">
		        <h3 id="editarRecursoModalLabel" class="displayInLine">
		            <spring:message code="update"/>&nbsp;<spring:message code="recurso"/>
		        </h3>
		    </div>
		    <div class="modal-body">
		        <form name="editarRecursoForm" novalidate class="well form-horizontal">
		            <input type="hidden"
		                   required
		                   ng-model="recurso.id"
		                   name="id"
		                   value="{{recurso.id}}"/>
					
					<div class="form-group" ng-class="{'has-error': mostrarErrorValidacion && editarRecursoForm.descripcion.$error.required}">
						<label for="txtDescripcion" class="col-lg-3 control-label"><spring:message code="recursos.descripcion"/>:</label>
						<div class="col-lg-7">
							<input type="text"
						    	class="form-control"
						    	id="txtDescripcion"
                               	required
                               	autofocus
                               	ng-model="recurso.descripcion"
                               	name="descripcion"
                               	placeholder="<spring:message code='recursos.descripcion'/>"/>         
						</div>
						<div class="col-lg-2">
							<label>
                            	<span class="alert alert-danger help-block"
                                      ng-show="mostrarErrorValidacion && editarRecursoForm.descripcion.$error.required">
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
			<div class="modal-footer">
			 <input type="submit"
		     	class="btn btn-primary"
		        ng-click="updateRecurso(editarRecursoForm);"
		        value='<spring:message code="update"/>'/>
		     <button class="btn btn-default"
		     	data-dismiss="modal"
		        ng-click="exit('#editarRecursoDialog');"
		        aria-hidden="true">
		        <spring:message code="cancel"/>
		     </button>
		</div>
		</div>
	</div>				    
</div>

<div id="eliminarRecursoDialog"
    class="modal fade"
    role="dialog"
    aria-hidden="true">
     
	<div class="modal-dialog">
   		<div class="modal-content">
    		<div class="modal-header">
    			<h4 id="eliminarRecursoDialogLabel" class="displayInLine">
            		<spring:message code="delete"/>&nbsp;<spring:message code="recurso"/>
        		</h4>
   			</div>
    	
			<div class="modal-body">
				<p><spring:message code="delete.confirm"/>:&nbsp;{{recurso.descripcion}}?</p>
				
				<span class="alert alert-danger"
		          ng-show="errorSubmit">
		        	<spring:message code="request.error"/>
		    	</span>
    			<span class="alert alert-danger"
          		ng-show="errorAccesoIlegal">
        			<spring:message code="request.illegal.access"/>
    			</span>
			</div>
			
			<div class="modal-footer">
			 	<form name="eliminarRecursoForm" novalidate>
            	<input type="submit"
                   class="btn btn-primary"
                   ng-click="eliminarRecurso();"
                   value='<spring:message code="delete"/>'/>
            	<button class="btn btn-default"
                    data-dismiss="modal"
                    ng-click="exit('#eliminarRecursoDialog');"
                    aria-hidden="true">
                	<spring:message code="cancel"/>
            	</button>
       			 </form>
			</div>    	
		</div>
    </div>
</div>

<!-- DIV con el dialogo de busqueda -->
<div id="buscarRecursosDialog"
     class="modal fade"
     role="dialog"
     aria-hidden="true">
    
    <div class="modal-dialog  modal-lg">
   		<div class="modal-content"> 
			<div class="modal-header">
        		<h3 id="buscarRecursosModalLabel" class="displayInLine">
            		<spring:message code="search"/>
        		</h3>
    		</div>
    	
    	    <!--     	DIV contenedor de los parametros de busqueda y los mensajes de error correspondientes -->
    		<div class="modal-body">
        		<form name="buscarRecursosForm" novalidate class="form-horizontal">
        			<!--         			DIV con la parte de busqueda -->
        			<div class="form-group" ng-class="{'has-error': mostrarErrorValidacion && buscarRecursosForm.filtroDescripcion.$error.required}">
        				<label for="txtDescripcion" class="col-lg-3 control-label"><spring:message code="search.for"/>:</label>
            			
	            		<div class="col-lg-7">
		                    	<input type="text"
		                    		class="form-control"
						    		id="txtDescripcion"	
		                           	autofocus
		                           	required
		                           	ng-model="filtroDescripcion"
		                           	name="filtroDescripcion"
		                           	placeholder="<spring:message code='recursos.descripcion'/> "/>
	                	</div>
	                	<div class="col-lg-2">
	                    	<label class="displayInLine">
	                        	<span class="alert alert-danger help-block"
	                              		ng-show="mostrarErrorValidacion && buscarRecursosForm.filtroDescripcion.$error.required">
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
                   		ng-click="buscarRecursos(buscarRecursosForm, false);"
                   		value='<spring:message code="search"/>'
                    	/>
            		<button class="btn btn-default"
                    	data-dismiss="modal"
                    	ng-click="exit('#buscarRecursosDialog');"
                    	aria-hidden="true">
                		<spring:message code="cancel"/>
            		</button>
    		</div>
    	</div>
    </div>			
 </div>

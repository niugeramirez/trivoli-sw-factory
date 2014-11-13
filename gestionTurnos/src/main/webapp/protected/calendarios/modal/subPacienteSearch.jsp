<%@ page contentType="text/html;charset=UTF-8" language="java" %>
<%@ taglib prefix="spring" uri="http://www.springframework.org/tags" %>

    <div id="pacienteSearchParameters"  class="modal-dialog  modal-lg">	    
   		<div class="modal-content"> 
			<div class="modal-header">
        		<h3 class="displayInLine">
            		<spring:message code="search"/>
        		</h3>
    		</div>
			
			<!--     	DIV contenedor de los parametros de busqueda y los mensajes de error correspondientes -->
    		<div class="modal-body">
        		<form name="busquedaForm" novalidate class="form-horizontal">
					
					<!--         			DIV con la parte de busqueda -->
        			<div class="form-group" ng-class="{'has-error': mostrarErrorValidacion }">
        				<label for="txtNombre" class="col-lg-3 control-label"><spring:message code="pacientes.DNI"/>:</label>            			
	            		<div class="col-lg-7">
		                    	<input type="text"
		                    		class="form-control"
						    		id="txtNombre"	
		                           	autofocus		                           	
		                           	ng-model="filtroPaciente.dni"
		                           	name="filtroPaciente.dni"
		                           	placeholder="<spring:message code='pacientes.DNI'/> "/>
	                	</div>
	                	        			
        				<label class="col-lg-3 control-label"><spring:message code="pacientes.apellido"/>:</label>            			
	            		<div class="col-lg-7">
		                    	<input type="text"
		                    		class="form-control"						    			
		                           	autofocus		                           	
		                           	ng-model="filtroPaciente.apellido"
		                           	name="filtroPaciente.apellido"
		                           	placeholder="<spring:message code='pacientes.apellido'/> "/>
	                	</div>    

        				<label class="col-lg-3 control-label"><spring:message code="pacientes.nombre"/>:</label>            			
	            		<div class="col-lg-7">
		                    	<input type="text"
		                    		class="form-control"						    			
		                           	autofocus		                           	
		                           	ng-model="filtroPaciente.nombre"
		                           	name="filtroPaciente.nombre"
		                           	placeholder="<spring:message code='pacientes.nombre'/> "/>
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
                 <a href="#pacienteQuickEditCreate"
                    ng-click="quickEditCreatePaciente(registro,'create');"
                    role="button"
                    title="<spring:message code="create"/>&nbsp;<spring:message code="paciente"/>"
                    class="btn btn-primary" data-toggle="modal">
                     <span class="glyphicon glyphicon-plus"></span>
                 </a>                
    			<input type="submit"
                   		class="btn btn-primary"
                   		ng-click="buscarPacientes();"
                   		title="<spring:message code="search"/>&nbsp;<spring:message code="paciente"/>"
                   		value='<spring:message code="search"/>'
                    	/>
            		<button class="btn btn-default"
                    	data-dismiss="modal"
                    	ng-click="exit('#pacienteAsign');"
                    	aria-hidden="true">
                		<spring:message code="cancel"/>
            		</button>
    		</div>
    	</div>
    </div>	
<%@ page contentType="text/html;charset=UTF-8" language="java" %>
<%@ taglib prefix="spring" uri="http://www.springframework.org/tags" %>    
    
<!-- DIV con el dialogo de confirmacion de la asignacion del turno -->
<div id="confirmTurnoAlta"
    class="modal fade"
    role="dialog"
    aria-hidden="true">
    
    <!-- 		DIV del  diálogo para el AJAX request -->
	<div id="loadingModalPracticas" class="modal hide fade in centering" role="dialog" aria-hidden="true">
           <div id="divLoadingIcon" class="text-center">
               <div class="glyphicon glyphicon-align-center loading"></div>
           </div>
    </div>
         
	<div class="modal-dialog">
   		<div class="modal-content">
    		<div class="modal-header">
    			<h4 class="displayInLine">
            		<spring:message code="turnos.confirm"/>&nbsp;
        		</h4>
   			</div>
    	
			<div class="modal-body">
				<p><label class="col-lg-3 control-label"><spring:message code="paciente"			/>:</label>{{pacienteActual.apellido}},{{pacienteActual.nombre}}		</p>
				<p><label class="col-lg-3 control-label"><spring:message code="recurso"			/>:</label>{{recursoActual.descripcion}}									</p>
				<p><label class="col-lg-3 control-label"><spring:message code="turnos.fecha"		/>:</label>{{calendarioActual.fechaHoraInicio | date:'dd/MM/yyyy'}}		</p>
				<p><label class="col-lg-3 control-label"><spring:message code="turnos.hora"		/>:</label>{{calendarioActual.fechaHoraInicio | date:'hh:mm'}} 				</p>
				<p><label class="col-lg-3 control-label"><spring:message code="practica"			/>:</label>
					<div class="col-lg-5"> 
						    <select class="form-control" 
		                 			ng-model="practicaActual" 
		                 			ng-options="pract.titulo  for pract in practicas" 
		                 			required 
		                 			name="practica"
		                 			>
		                 	</select>
		             </div></p>
   
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
			
			<!--     		DIV con los botones de aceptar y cancelar -->
			<div class="modal-footer">
			 	<form name="eliminarObraSocialForm" novalidate>
            	<input type="submit"
                   class="btn btn-primary"
                   ng-click="altaTurno();"
                   value='<spring:message code="accept"/>'/>
                   
            	<button class="btn btn-default"
                    data-dismiss="modal"
                    ng-click="exitQuickEditCreate();"
                    aria-hidden="true">
                	<spring:message code="cancel"/>
            	</button>
       			 </form>
			</div>    	
		</div>
    </div>
</div>
    
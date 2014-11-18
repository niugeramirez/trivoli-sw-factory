/**
 * 
 */
package ar.com.trivoli.gestionturnos.calendario.controller;

import java.util.Date;
import java.util.List;
import java.util.Locale;











import org.springframework.beans.factory.annotation.Autowired;
import org.springframework.format.annotation.DateTimeFormat;
import org.springframework.http.HttpStatus;
import org.springframework.http.ResponseEntity;
import org.springframework.stereotype.Controller;
import org.springframework.web.bind.annotation.PathVariable;
import org.springframework.web.bind.annotation.RequestBody;
import org.springframework.web.bind.annotation.RequestMapping;
import org.springframework.web.bind.annotation.RequestMethod;
import org.springframework.web.bind.annotation.RequestParam;
import org.springframework.web.servlet.ModelAndView;

import ar.com.trivoli.gestionturnos.calendario.model.Calendario;
import ar.com.trivoli.gestionturnos.calendario.model.CalendarioDTO;
import ar.com.trivoli.gestionturnos.calendario.service.CalendarioService;
import ar.com.trivoli.gestionturnos.common.controller.ControllerBase;
import ar.com.trivoli.gestionturnos.common.model.ListaEntidadDTO;
import ar.com.trivoli.gestionturnos.obrasocial.model.ObraSocial;
import ar.com.trivoli.gestionturnos.obrasocial.service.ObraSocialService;
import ar.com.trivoli.gestionturnos.paciente.model.Paciente;
import ar.com.trivoli.gestionturnos.paciente.service.PacienteService;
import ar.com.trivoli.gestionturnos.practica.model.Practica;
import ar.com.trivoli.gestionturnos.practica.service.PracticaService;
import ar.com.trivoli.gestionturnos.recurso.model.Recurso;
import ar.com.trivoli.gestionturnos.recurso.service.RecursoService;
import ar.com.trivoli.gestionturnos.turno.model.Turno;
import ar.com.trivoli.gestionturnos.turno.service.TurnoService;




/**
 * @author ramirez
 *
 */

@Controller

@RequestMapping(value = "/protected/calendarios")
public class CalendarioController extends ControllerBase<Calendario> {

	@Autowired
	private CalendarioService calendarioService;
	
	@Autowired
	private RecursoService recursoService;
	
	@Autowired
	private PacienteService	pacienteService;
	
	@Autowired
	private ObraSocialService obraSocialService;

	@Autowired
	private PracticaService practicaService;
	
	@Autowired
	private TurnoService turnoService;
	
	
	//TODO configurar motor de testing
	
	/************************************************************************************************************************************************************************/
	@RequestMapping(method = RequestMethod.GET)
	public ModelAndView welcome() {
		return new ModelAndView("admCalendarios");
	}
	/************************************************************************************************************************************************************************/
	@RequestMapping(method = RequestMethod.GET, produces = "application/json")
	public ResponseEntity<?> listarTodosLosRecursos(@RequestParam int nroPagina,Locale locale) {

		
		// Se recuperan todos los registros
		List<Recurso> listaRecursos = recursoService.recuperarTodos();
		ListaEntidadDTO<Recurso> listaRecursosDTO = new ListaEntidadDTO<Recurso>(1,listaRecursos.size(),listaRecursos);

		// Se arma la Respuesta HTTP
		return new ResponseEntity<ListaEntidadDTO<Recurso>>(listaRecursosDTO,HttpStatus.OK);
	}
	/************************************************************************************************************************************************************************/
	@RequestMapping(value = "/{idRecurso}", method = RequestMethod.GET, produces = "application/json")
	public ResponseEntity<?> listAllPorRecurso(	@PathVariable("idRecurso") int idRecurso
												,@RequestParam @DateTimeFormat(pattern="dd-MM-yyyy") Date  fechaTurnos
												,@RequestParam int nroPagina,Locale locale) {
			
		// Se recuperan todos los registros
		ListaEntidadDTO<CalendarioDTO> listaCalendarios = calendarioService.recuperarCalendariosPorRecursoYFecha(idRecurso,fechaTurnos);

		// Se arma la Respuesta HTTP
		return new ResponseEntity<ListaEntidadDTO<CalendarioDTO>>(listaCalendarios,HttpStatus.OK);
	}
	/************************************************************************************************************************************************************************/
	@RequestMapping(value = "/{idRecurso}/{idCalendario}", method = RequestMethod.GET, produces = "application/json")
	public ResponseEntity<?> buscarTurnos(
			@PathVariable("idCalendario") int idCalendario,
			@RequestParam(required = false, defaultValue = DEFAULT_PAGE_DISPLAYED_TO_USER) int nroPagina,
			Locale locale) {		

		ListaEntidadDTO<Turno> listaTurnos =  calendarioService.recuperarTurnos(idCalendario);

		return new ResponseEntity<ListaEntidadDTO<Turno>>(listaTurnos
															,HttpStatus.OK);
	}

	/************************************************************************************************************************************************************************/
	@RequestMapping(value = "/turno", method = RequestMethod.PUT, produces = "application/json")
	public ResponseEntity<?> updateTurno(

			@RequestBody Turno turno,			
			@RequestParam(required = false, defaultValue = DEFAULT_PAGE_DISPLAYED_TO_USER) int nroPagina,
			Locale locale) {		

		turnoService.guardar(turno);
		
		ListaEntidadDTO<Turno> listaTurnos =  calendarioService.recuperarTurnos(turno.getCalendario().getId());

		return new ResponseEntity<ListaEntidadDTO<Turno>>(listaTurnos
															,HttpStatus.OK);
	}
	
	/************************************************************************************************************************************************************************/
	//TODO unificar las distintas URLs en todos los controles por algo más prolijo y consistente
	@RequestMapping(value = "/pacientes", method = RequestMethod.GET, produces = "application/json")
	public ResponseEntity<?> buscarPacientes(	@RequestParam(required = false) String filtroDNI,
												@RequestParam(required = false) String filtroApellido,
												@RequestParam(required = false) String filtroNombre,
												@RequestParam(required = false, defaultValue = DEFAULT_PAGE_DISPLAYED_TO_USER) int nroPagina,
												Locale locale) 
	{	
		
		ListaEntidadDTO<Paciente> listaPacientes = pacienteService.recuperarPorComienzoDniApellidoNombreLike(nroPagina 
																											,registrosPorPagina
																											,filtroDNI
																											,filtroApellido
																											,filtroNombre
																											);
		

		return new ResponseEntity<ListaEntidadDTO<Paciente>>(listaPacientes,HttpStatus.OK);
	}
	/************************************************************************************************************************************************************************/
	@RequestMapping(value = "/pacientes", method = RequestMethod.PUT, produces = "application/json")
	public ResponseEntity<?> updatePaciente(
			@RequestBody Paciente paciente,
			@RequestParam(required = false) String filtroDNI,
			@RequestParam(required = false) String filtroApellido,
			@RequestParam(required = false) String filtroNombre,
			@RequestParam(required = false, defaultValue = DEFAULT_PAGE_DISPLAYED_TO_USER) int nroPagina,
			Locale locale)  
	{				

		try {
			pacienteService.guardar(paciente);
		} catch (Exception e) {
			return new ResponseEntity<String>(e.getMessage(),HttpStatus.BAD_REQUEST);
		}

		ListaEntidadDTO<Paciente> listaPacientes = pacienteService.recuperarPorComienzoDniApellidoNombreLike(nroPagina 
				,registrosPorPagina
				,filtroDNI
				,filtroApellido
				,filtroNombre
				);


		return new ResponseEntity<ListaEntidadDTO<Paciente>>(listaPacientes,HttpStatus.OK);
	}	
	
	/************************************************************************************************************************************************************************/
	@RequestMapping(value = "/obrasSociales", method = RequestMethod.GET, produces = "application/json")
	public ResponseEntity<?> buscarObrasSociales(@RequestParam int nroPagina,Locale locale) {

		
		// Se recuperan todos los registros
		List<ObraSocial> listaObraSocial = obraSocialService.recuperarTodos();
		ListaEntidadDTO<ObraSocial> listaRecursosDTO = new ListaEntidadDTO<ObraSocial>(1,listaObraSocial.size(),listaObraSocial);

		// Se arma la Respuesta HTTP
		return new ResponseEntity<ListaEntidadDTO<ObraSocial>>(listaRecursosDTO,HttpStatus.OK);
	}	
	/************************************************************************************************************************************************************************/
	@RequestMapping(value = "/practicas", method = RequestMethod.GET, produces = "application/json")
	public ResponseEntity<?> buscarPracticas(@RequestParam int nroPagina,Locale locale) {

		
		// Se recuperan todos los registros
		List<Practica> listaPractica = practicaService.recuperarTodos();
		ListaEntidadDTO<Practica> listaRecursosDTO = new ListaEntidadDTO<Practica>(1,listaPractica.size(),listaPractica);

		// Se arma la Respuesta HTTP
		return new ResponseEntity<ListaEntidadDTO<Practica>>(listaRecursosDTO,HttpStatus.OK);
	}	
	/************************************************************************************************************************************************************************/

}

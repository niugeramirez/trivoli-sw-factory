/**
 * 
 */
package ar.com.trivoli.gestionturnos.calendario.controller;

import java.util.Date;
import java.util.List;
import java.util.Locale;

import org.springframework.beans.factory.annotation.Autowired;
import org.springframework.http.HttpStatus;
import org.springframework.http.ResponseEntity;
import org.springframework.stereotype.Controller;
import org.springframework.web.bind.annotation.PathVariable;
import org.springframework.web.bind.annotation.RequestMapping;
import org.springframework.web.bind.annotation.RequestMethod;
import org.springframework.web.bind.annotation.RequestParam;
import org.springframework.web.servlet.ModelAndView;

import ar.com.trivoli.gestionturnos.calendario.model.Calendario;
import ar.com.trivoli.gestionturnos.calendario.model.CalendarioDTO;
import ar.com.trivoli.gestionturnos.calendario.service.CalendarioService;
import ar.com.trivoli.gestionturnos.common.controller.ControllerBase;
import ar.com.trivoli.gestionturnos.common.model.ListaEntidadDTO;
import ar.com.trivoli.gestionturnos.recurso.model.Recurso;
import ar.com.trivoli.gestionturnos.recurso.service.RecursoService;
import ar.com.trivoli.gestionturnos.turno.model.Turno;




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
	
	//TODO configurar motor de testing
	
	/************************************************************************************************************************************************************************/
	@RequestMapping(method = RequestMethod.GET)
	public ModelAndView welcome() {
		return new ModelAndView("admCalendarios");
	}
	/************************************************************************************************************************************************************************/
	@RequestMapping(method = RequestMethod.GET, produces = "application/json")
	public ResponseEntity<?> listarTodosLosRecursos(@RequestParam int nroPagina,Locale locale) {

		System.out.println("listando todos los recursos");
		// Se recuperan todos los registros
		List<Recurso> listaRecursos = recursoService.recuperarTodos();
		ListaEntidadDTO<Recurso> listaRecursosDTO = new ListaEntidadDTO<Recurso>(1,listaRecursos.size(),listaRecursos);

		// Se arma la Respuesta HTTP
		return new ResponseEntity<ListaEntidadDTO<Recurso>>(listaRecursosDTO,HttpStatus.OK);
	}
	/************************************************************************************************************************************************************************/
	@RequestMapping(value = "/{idRecurso}", method = RequestMethod.GET, produces = "application/json")
	public ResponseEntity<?> listAllPorRecurso(	@PathVariable("idRecurso") int idRecurso
												,@RequestParam Date fechaTurnos
												,@RequestParam int nroPagina,Locale locale) {

		System.out.println("entro el list all por recurso");
		System.out.println("idRecurso "+idRecurso);
		System.out.println("fechaTurnos "+fechaTurnos);
		System.out.println("fechaTurnos to string "+fechaTurnos.toString() );

		Date fecha1 = new Date ();		   
		System.out.println(fecha1.toString());
		
		// Se recuperan todos los registros
		ListaEntidadDTO<CalendarioDTO> listaCalendarios = calendarioService.recuperarCalendariosPorRecusro(idRecurso);

		// Se arma la Respuesta HTTP
		return new ResponseEntity<ListaEntidadDTO<CalendarioDTO>>(listaCalendarios,HttpStatus.OK);
	}
	/************************************************************************************************************************************************************************/
	@RequestMapping(value = "/{idRecurso}/{idCalendario}", method = RequestMethod.GET, produces = "application/json")
	public ResponseEntity<?> buscarTurnos(
			@PathVariable("idCalendario") int idCalendario,
			@RequestParam(required = false, defaultValue = DEFAULT_PAGE_DISPLAYED_TO_USER) int nroPagina,
			Locale locale) {
		System.out.println("entro el search");
		System.out.println("idCalendario "+ idCalendario);
		Date fecha1 = new Date ();
		System.out.println(fecha1.toString());

		ListaEntidadDTO<Turno> listaTurnos =  calendarioService.recuperarTurnos(idCalendario);

		return new ResponseEntity<ListaEntidadDTO<Turno>>(listaTurnos
															,HttpStatus.OK);
	}
	/************************************************************************************************************************************************************************/
	
}

/**
 * 
 */
package ar.com.trivoli.gestionturnos.calendario.controller;

import java.util.Locale;

import org.springframework.beans.factory.annotation.Autowired;
import org.springframework.http.HttpStatus;
import org.springframework.http.ResponseEntity;
import org.springframework.stereotype.Controller;
import org.springframework.web.bind.annotation.RequestMapping;
import org.springframework.web.bind.annotation.RequestMethod;
import org.springframework.web.bind.annotation.RequestParam;
import org.springframework.web.servlet.ModelAndView;

import ar.com.trivoli.gestionturnos.calendario.model.Calendario;
import ar.com.trivoli.gestionturnos.calendario.model.CalendarioDTO;
import ar.com.trivoli.gestionturnos.calendario.service.CalendarioService;
import ar.com.trivoli.gestionturnos.common.controller.ControllerBase;
import ar.com.trivoli.gestionturnos.common.model.ListaEntidadDTO;




/**
 * @author ramirez
 *
 */

@Controller

@RequestMapping(value = "/protected/calendario")
public class CalendarioController extends ControllerBase<Calendario> {

	@Autowired
	private CalendarioService calendarioService;
	/************************************************************************************************************************************************************************/
	@RequestMapping(method = RequestMethod.GET)
	public ModelAndView welcome() {
		return new ModelAndView("admCalendario");
	}	
	/************************************************************************************************************************************************************************/
	/**
	 * Método que se invoca al recibir un GET que espera como respuesta un
	 * objeto JSON
	 * 
	 * @param nroPagina
	 *            Numero de Pagina solicitada
	 * @param locale
	 *            Informacion de Localizacion
	 * @return HTTP Response
	 */	
	@RequestMapping(method = RequestMethod.GET, produces = "application/json")
	public ResponseEntity<?> listAll(@RequestParam int nroPagina, Locale locale) {

		// Se recuperan todos los registros
		ListaEntidadDTO<CalendarioDTO> listaCalendarios = calendarioService.recuperarTodos();

		// Se arma la Respuesta HTTP
		return new ResponseEntity<ListaEntidadDTO<CalendarioDTO>>(listaCalendarios,HttpStatus.OK);
	}
	/************************************************************************************************************************************************************************/	
}

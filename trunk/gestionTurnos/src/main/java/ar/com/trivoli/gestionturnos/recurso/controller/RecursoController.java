/**
 * 
 */
package ar.com.trivoli.gestionturnos.recurso.controller;

import java.util.Locale;

import org.apache.commons.lang.StringUtils;
import org.springframework.beans.factory.annotation.Autowired;
import org.springframework.http.HttpStatus;
import org.springframework.http.ResponseEntity;
import org.springframework.security.access.AccessDeniedException;
import org.springframework.stereotype.Controller;
import org.springframework.web.bind.annotation.ModelAttribute;
import org.springframework.web.bind.annotation.PathVariable;
import org.springframework.web.bind.annotation.RequestBody;
import org.springframework.web.bind.annotation.RequestMapping;
import org.springframework.web.bind.annotation.RequestMethod;
import org.springframework.web.bind.annotation.RequestParam;
import org.springframework.web.servlet.ModelAndView;

import ar.com.trivoli.gestionturnos.common.controller.ControllerBase;
import ar.com.trivoli.gestionturnos.common.model.ListaEntidadDTO;
import ar.com.trivoli.gestionturnos.recurso.model.Recurso;
import ar.com.trivoli.gestionturnos.recurso.service.RecursoService;

/**
 * @author posadas
 * 
 *         Controlador Pagina de Administración de Recursos
 */
@Controller
@RequestMapping(value = "/protected/recursos")
public class RecursoController extends ControllerBase<Recurso> {

	@Autowired
	private RecursoService recursoService;

	/**
	 * Maxima Cantidad de Registros por Pagina
	 */

	private ResponseEntity<?> buscarRecursos(String filtroDescripcion,
			int nroPagina, Locale locale, String actionMessageKey) {
		ListaEntidadDTO<Recurso> listaRecursos = recursoService
				.buscarRecursosPorDescripcion(nroPagina, registrosPorPagina,
						filtroDescripcion);

		if (!StringUtils.isEmpty(actionMessageKey)) {
			agregarMensajeAccion(listaRecursos, locale, actionMessageKey, null);
		}

		Object[] args = { filtroDescripcion };

		agregarMensajeBusqueda(listaRecursos, locale,
				"message.search.for.active", args);

		return new ResponseEntity<ListaEntidadDTO<Recurso>>(listaRecursos,
				HttpStatus.OK);
	}

	@RequestMapping(method = RequestMethod.POST, produces = "application/json")
	public ResponseEntity<?> create(
			@ModelAttribute("recurso") Recurso recurso,
			@RequestParam(required = false) String filtroDescripcion,
			@RequestParam(required = false, defaultValue = DEFAULT_PAGE_DISPLAYED_TO_USER) int nroPagina,
			Locale locale) {
		recursoService.guardar(recurso);

		if (existeBusquedaActiva(filtroDescripcion)) {
			return buscarRecursos(filtroDescripcion, nroPagina, locale,
					"message.create.success");
		}

		// Se recuperan todos los Recursos
		ListaEntidadDTO<Recurso> listaRecursos = recursoService.recuperarTodos(
				nroPagina, registrosPorPagina);

		agregarMensajeAccion(listaRecursos, locale, "message.create.success",
				null);

		return new ResponseEntity<ListaEntidadDTO<Recurso>>(listaRecursos,
				HttpStatus.OK);
	}

	@RequestMapping(value = "/{recursoId}", method = RequestMethod.DELETE, produces = "application/json")
	public ResponseEntity<?> delete(
			@PathVariable("recursoId") int recursoId,
			@RequestParam(required = false) String filtroDescripcion,
			@RequestParam(required = false, defaultValue = DEFAULT_PAGE_DISPLAYED_TO_USER) int nroPagina,
			Locale locale) {

		try {
			recursoService.delete(recursoId);
		} catch (AccessDeniedException e) {
			return new ResponseEntity<Object>(HttpStatus.FORBIDDEN);
		}

		if (existeBusquedaActiva(filtroDescripcion)) {
			return buscarRecursos(filtroDescripcion, nroPagina, locale,
					"message.delete.success");
		}

		// Se recuperan todos los Recursos
		ListaEntidadDTO<Recurso> listaRecursos = recursoService.recuperarTodos(
				nroPagina, registrosPorPagina);

		agregarMensajeAccion(listaRecursos, locale, "message.delete.success",
				null);

		return new ResponseEntity<ListaEntidadDTO<Recurso>>(listaRecursos,
				HttpStatus.OK);
	}

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
		// Se recuperan todos los Recursos
		ListaEntidadDTO<Recurso> listaRecursos = recursoService.recuperarTodos(
				nroPagina, registrosPorPagina);

		// Se arma la Respuesta HTTP
		return new ResponseEntity<ListaEntidadDTO<Recurso>>(listaRecursos,
				HttpStatus.OK);
	}

	@RequestMapping(value = "/{filtroDescripcion}", method = RequestMethod.GET, produces = "application/json")
	public ResponseEntity<?> search(
			@PathVariable("filtroDescripcion") String filtroDescripcion,
			@RequestParam(required = false, defaultValue = DEFAULT_PAGE_DISPLAYED_TO_USER) int nroPagina,
			Locale locale) {
		return buscarRecursos(filtroDescripcion, nroPagina, locale, null);
	}

	@RequestMapping(value = "/{id}", method = RequestMethod.PUT, produces = "application/json")
	public ResponseEntity<?> update(
			@PathVariable("id") int contactId,
			@RequestBody Recurso recurso,
			@RequestParam(required = false) String filtroDescripcion,
			@RequestParam(required = false, defaultValue = DEFAULT_PAGE_DISPLAYED_TO_USER) int nroPagina,
			Locale locale) {
		if (contactId != recurso.getId()) {
			return new ResponseEntity<String>("Solicitud Incorrecta",
					HttpStatus.BAD_REQUEST);
		}

		recursoService.guardar(recurso);

		if (existeBusquedaActiva(filtroDescripcion)) {
			return buscarRecursos(filtroDescripcion, nroPagina, locale,
					"message.update.success");
		}

		// Se recuperan todos los Recursos
		ListaEntidadDTO<Recurso> listaRecursos = recursoService.recuperarTodos(
				nroPagina, registrosPorPagina);

		agregarMensajeAccion(listaRecursos, locale, "message.update.success",
				null);

		return new ResponseEntity<ListaEntidadDTO<Recurso>>(listaRecursos,
				HttpStatus.OK);
	}

	/**
	 * Método que se invoca al recibir un GET desde el Front
	 * 
	 * @return
	 */
	@RequestMapping(method = RequestMethod.GET)
	public ModelAndView welcome() {
		return new ModelAndView("admRecursos");
	}
}

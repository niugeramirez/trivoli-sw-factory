/**
 * 
 */
package ar.com.trivoli.gestionturnos.common.controller;

import java.util.Locale;

import org.apache.commons.lang.StringUtils;
import org.springframework.beans.factory.annotation.Autowired;
import org.springframework.beans.factory.annotation.Value;
import org.springframework.context.MessageSource;

import ar.com.trivoli.gestionturnos.common.model.ListaEntidadDTO;

/**
 * @author posadas
 * 
 *         Controller Base que provee caracteristicas comunes como la maxima
 *         cantidad de Registros por Pagina, el acceso al Repositorio de
 *         Mensajes y métodos para pasar Mensajes al Usuario
 * 
 */
public class ControllerBase<tipoId> {
	protected static final String DEFAULT_PAGE_DISPLAYED_TO_USER = "0";

	/**
	 * Maxima Cantidad de Registros por Pagina
	 */
	@Value("5")
	protected int registrosPorPagina;

	@Autowired
	protected MessageSource messageSource;

	protected ListaEntidadDTO<tipoId> agregarMensajeAccion(
			ListaEntidadDTO<tipoId> listaEntidad, Locale locale,
			String claveMensaje, Object[] args) {
		if (StringUtils.isEmpty(claveMensaje)) {
			return listaEntidad;
		}

		listaEntidad.setMensajeAccion(messageSource.getMessage(claveMensaje,
				args, null, locale));

		return listaEntidad;
	}

	protected ListaEntidadDTO<tipoId> agregarMensajeBusqueda(
			ListaEntidadDTO<tipoId> listaEntidad, Locale locale,
			String claveMensaje, Object[] args) {
		if (StringUtils.isEmpty(claveMensaje)) {
			return listaEntidad;
		}

		listaEntidad.setMensajeBusqueda(messageSource.getMessage(claveMensaje,
				args, null, locale));

		return listaEntidad;
	}

	protected boolean existeBusquedaActiva(String filtroDescripcion) {
		return !StringUtils.isEmpty(filtroDescripcion);
	}
}

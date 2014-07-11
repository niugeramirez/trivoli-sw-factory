/**
 * 
 */
package ar.com.trivoli.gestionturnos.common.model;

import java.util.List;

/**
 * @author posadas
 * 
 *         DTO que permite transferir lista de Objetos entre las capas de la
 *         Aplicacion, con informacion de Paginado, etc.
 * 
 */
public class ListaEntidadDTO<tipoId> {
	private int cantPaginas;
	private long totalRegistros;

	private String mensajeAccion;
	private String mensajeBusqueda;

	private List<tipoId> registros;

	public ListaEntidadDTO() {
	}

	public ListaEntidadDTO(int cantPaginas, long totalRegistros,
			List<tipoId> registros) {
		this.cantPaginas = cantPaginas;
		this.registros = registros;
		this.totalRegistros = totalRegistros;
	}

	public int getCantPaginas() {
		return cantPaginas;
	}

	public String getMensajeAccion() {
		return mensajeAccion;
	}

	public String getMensajeBusqueda() {
		return mensajeBusqueda;
	}

	public List<tipoId> getRegistros() {
		return registros;
	}

	public long getTotalRegistros() {
		return totalRegistros;
	}

	public void setCantPaginas(int cantPaginas) {
		this.cantPaginas = cantPaginas;
	}

	public void setMensajeAccion(String mensajeAccion) {
		this.mensajeAccion = mensajeAccion;
	}

	public void setMensajeBusqueda(String mensajeBusqueda) {
		this.mensajeBusqueda = mensajeBusqueda;
	}

	public void setRegistros(List<tipoId> registros) {
		this.registros = registros;
	}

	public void setTotalRegistros(long totalRegistros) {
		this.totalRegistros = totalRegistros;
	}

}

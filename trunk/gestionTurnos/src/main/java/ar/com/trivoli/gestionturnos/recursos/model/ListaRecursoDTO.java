/**
 * 
 */
package ar.com.trivoli.gestionturnos.recursos.model;

import java.util.List;

/**
 * @author posadas
 * 
 *         DTO que representa una Lista de Recurso con informacion de Paginado y
 *         Mensajes al Usuario Final
 */
public class ListaRecursoDTO {
	private int cantPaginas;
	private long totalRegistros;

	private String mensajeAccion;
	private String mensajeBusqueda;

	private List<Recurso> registros;

	public ListaRecursoDTO() {
	}

	public ListaRecursoDTO(int cantPaginas, long totalRegistros,
			List<Recurso> recursos) {
		this.cantPaginas = cantPaginas;
		this.registros = recursos;
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

	public List<Recurso> getRegistros() {
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

	public void setRegistros(List<Recurso> registros) {
		this.registros = registros;
	}

	public void setTotalRegistros(long totalRegistros) {
		this.totalRegistros = totalRegistros;
	}

}

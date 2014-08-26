/**
 * 
 */
package ar.com.trivoli.gestionturnos.obrasocial.service;

import org.springframework.beans.factory.annotation.Autowired;
import org.springframework.data.domain.Page;
import org.springframework.data.domain.PageRequest;
import org.springframework.data.domain.Sort;
import org.springframework.stereotype.Service;
import org.springframework.transaction.annotation.Transactional;

import ar.com.trivoli.gestionturnos.common.model.ListaEntidadDTO;
import ar.com.trivoli.gestionturnos.obrasocial.model.ObraSocial;
import ar.com.trivoli.gestionturnos.obrasocial.repository.IObraSocialRepository;



/**
 * @author ramirez
 *
 *			Servicio de Obras Sociales
 */

@Service
@Transactional
public class ObraSocialService {

	/**
	 * 
	 */
	
	@Autowired
	private IObraSocialRepository obrasocialRepository;

	/************************************************************************************************************************************************************************/
	private Sort ordenPredeterminado() {
		return new Sort(Sort.Direction.ASC, "nombre");
	}	
	/************************************************************************************************************************************************************************/
	public void delete(int obraSocialId) {
		obrasocialRepository.delete(obraSocialId);
	}
	/************************************************************************************************************************************************************************/
	public void guardar(ObraSocial obraSocial) {
		obrasocialRepository.save(obraSocial);
	}
	/************************************************************************************************************************************************************************/	
	@Transactional(readOnly = true)
	public ListaEntidadDTO<ObraSocial> buscarObrasSocialesPorNombre(	int nroPagina,
																		int registrosPorPagina, 
																		String descripcion) {
		PageRequest pageRequest = new PageRequest(	nroPagina,
													registrosPorPagina, 
													ordenPredeterminado());

		Page<ObraSocial> resultado = obrasocialRepository.findByNombreLike(	pageRequest,
																			"%" + descripcion + "%");

		// Se determina si la pagina requerida es posterior a la Ultima pagina
		// disponible
		if (resultado.getTotalElements() > 0
				&& nroPagina > (resultado.getTotalPages() - 1)) {
			int ultimaPagina = resultado.getTotalPages() - 1;

			pageRequest = new PageRequest(ultimaPagina, registrosPorPagina,
					ordenPredeterminado());

			resultado = obrasocialRepository.findByNombreLike(	pageRequest,
																"%" + descripcion + "%");
		}

		return new ListaEntidadDTO<ObraSocial>(	resultado.getTotalPages(),
												resultado.getTotalElements(), 
												resultado.getContent());
	}	
	
	/************************************************************************************************************************************************************************/
	@Transactional(readOnly = true)
	public ListaEntidadDTO<ObraSocial> recuperarTodos(int nroPagina,
			int registrosPorPagina) {
		PageRequest pageRequest = new PageRequest(nroPagina,
				registrosPorPagina, ordenPredeterminado());

		Page<ObraSocial> resultado = obrasocialRepository.findAll(pageRequest);

		// Se determina si la pagina requerida es posterior a la Ultima pagina
		// disponible
		if (resultado.getTotalElements() > 0
				&& nroPagina > (resultado.getTotalPages() - 1)) {
			int ultimaPagina = resultado.getTotalPages() - 1;

			pageRequest = new PageRequest(ultimaPagina, registrosPorPagina,
					ordenPredeterminado());

			resultado = obrasocialRepository.findAll(pageRequest);
		}

		return new ListaEntidadDTO<ObraSocial>(resultado.getTotalPages(),
				resultado.getTotalElements(), resultado.getContent());
	}
	/************************************************************************************************************************************************************************/
}

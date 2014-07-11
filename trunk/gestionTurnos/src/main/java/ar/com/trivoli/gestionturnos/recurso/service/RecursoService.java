/**
 * 
 */
package ar.com.trivoli.gestionturnos.recurso.service;

import org.springframework.beans.factory.annotation.Autowired;
import org.springframework.data.domain.Page;
import org.springframework.data.domain.PageRequest;
import org.springframework.data.domain.Sort;
import org.springframework.stereotype.Service;
import org.springframework.transaction.annotation.Transactional;

import ar.com.trivoli.gestionturnos.common.model.ListaEntidadDTO;
import ar.com.trivoli.gestionturnos.recurso.model.Recurso;
import ar.com.trivoli.gestionturnos.recurso.repository.IRecursoRepository;

/**
 * @author posadas
 * 
 *         Servicio de Recursos
 */
@Service
@Transactional
public class RecursoService {

	@Autowired
	private IRecursoRepository recursoRepository;

	private Sort ordenPredeterminado() {
		return new Sort(Sort.Direction.ASC, "descripcion");
	}

	@Transactional(readOnly = true)
	public ListaEntidadDTO<Recurso> buscarRecursosPorDescripcion(int nroPagina,
			int registrosPorPagina, String descripcion) {
		PageRequest pageRequest = new PageRequest(nroPagina,
				registrosPorPagina, ordenPredeterminado());

		Page<Recurso> resultado = recursoRepository.findByDescripcionLike(
				pageRequest, "%" + descripcion + "%");

		// Se determina si la pagina requerida es posterior a la Ultima pagina
		// disponible
		if (resultado.getTotalElements() > 0
				&& nroPagina > (resultado.getTotalPages() - 1)) {
			int ultimaPagina = resultado.getTotalPages() - 1;

			pageRequest = new PageRequest(ultimaPagina, registrosPorPagina,
					ordenPredeterminado());

			resultado = recursoRepository.findByDescripcionLike(pageRequest,
					"%" + descripcion + "%");
		}

		return new ListaEntidadDTO<Recurso>(resultado.getTotalPages(),
				resultado.getTotalElements(), resultado.getContent());
	}

	public void delete(int recursoId) {
		recursoRepository.delete(recursoId);
	}

	public void guardar(Recurso recurso) {
		recursoRepository.save(recurso);
	}

	@Transactional(readOnly = true)
	public ListaEntidadDTO<Recurso> recuperarTodos(int nroPagina,
			int registrosPorPagina) {
		PageRequest pageRequest = new PageRequest(nroPagina,
				registrosPorPagina, ordenPredeterminado());

		Page<Recurso> resultado = recursoRepository.findAll(pageRequest);

		// Se determina si la pagina requerida es posterior a la Ultima pagina
		// disponible
		if (resultado.getTotalElements() > 0
				&& nroPagina > (resultado.getTotalPages() - 1)) {
			int ultimaPagina = resultado.getTotalPages() - 1;

			pageRequest = new PageRequest(ultimaPagina, registrosPorPagina,
					ordenPredeterminado());

			resultado = recursoRepository.findAll(pageRequest);
		}

		return new ListaEntidadDTO<Recurso>(resultado.getTotalPages(),
				resultado.getTotalElements(), resultado.getContent());
	}

}

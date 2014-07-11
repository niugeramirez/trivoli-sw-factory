/**
 * 
 */
package ar.com.trivoli.gestionturnos.modeloturno.service;

import org.springframework.beans.factory.annotation.Autowired;
import org.springframework.data.domain.Page;
import org.springframework.data.domain.PageRequest;
import org.springframework.data.domain.Sort;
import org.springframework.stereotype.Service;
import org.springframework.transaction.annotation.Transactional;

import ar.com.trivoli.gestionturnos.common.model.ListaEntidadDTO;
import ar.com.trivoli.gestionturnos.modeloturno.model.ModeloTurno;
import ar.com.trivoli.gestionturnos.modeloturno.repository.IModeloTurnoRepository;

/**
 * @author posadas
 * 
 *         Servicio de Modelos de Turnos
 */
@Service
@Transactional
public class ModeloTurnoService {
	@Autowired
	private IModeloTurnoRepository modeloTurnoRepository;

	private Sort ordenPredeterminado() {
		return new Sort(Sort.Direction.ASC, "descripcion");
	}

	@Transactional(readOnly = true)
	public ListaEntidadDTO<ModeloTurno> recuperarTodos(int nroPagina,
			int registrosPorPagina) {
		PageRequest pageRequest = new PageRequest(nroPagina,
				registrosPorPagina, ordenPredeterminado());

		Page<ModeloTurno> resultado = modeloTurnoRepository
				.findAll(pageRequest);

		// Se determina si la pagina requerida es posterior a la Ultima pagina
		// disponible
		if (resultado.getTotalElements() > 0
				&& nroPagina > (resultado.getTotalPages() - 1)) {
			int ultimaPagina = resultado.getTotalPages() - 1;

			pageRequest = new PageRequest(ultimaPagina, registrosPorPagina,
					ordenPredeterminado());

			resultado = modeloTurnoRepository.findAll(pageRequest);
		}

		return new ListaEntidadDTO<ModeloTurno>(resultado.getTotalPages(),
				resultado.getTotalElements(), resultado.getContent());
	}
}

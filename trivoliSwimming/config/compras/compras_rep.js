 function inicalizar(localdata_master,localdata_detail_all,localdata_detail_all_pago) {
            var source_master =
            {
                datafields: [
                    { name: 'id_compra' },
                    { name: 'nombre_proveedor' },
                    { name: 'fecha'  },
                    { name: 'monto_compra' },
                    { name: 'pagado' },
                    { name: 'saldo' },
					//,{ name: 'CustomerID' }
                ],
                localdata: localdata_master
            };
            var dataAdapter_master = new $.jqx.dataAdapter(source_master);
            $("#comprasGrid").jqxGrid(
            {
                width: 750,
                height: 250,
                source: dataAdapter_master,
                
                keyboardnavigation: false,
                columns: [
                    { text: 'Proveedor', datafield: 'nombre_proveedor', width: 250 },
                    { text: 'Fecha', datafield: 'fecha', cellsformat: 'd',width: 150 },
                    { text: 'Monto Compra', datafield: 'monto_compra', width: 110 },
                    { text: 'Pagado', datafield: 'pagado', width: 110 },
                    { text: 'Saldo', datafield: 'saldo', width: 110}
                ]
            });
            			
			
			/*****************************Detalle compras  prepare the data*****************************/
            var dataFields_detail = [
                        { name: 'descripcion_articulo' },
                        { name: 'cantidad' },
                        { name: 'precio_unitario' },
                        { name: 'subtotal' },
                        { name: 'id_compra' }
                    ];
            var source_detail =
            {
                datafields: dataFields_detail,
                localdata: localdata_detail_all
			};
            var dataAdapter_detail = new $.jqx.dataAdapter(source_detail);
            dataAdapter_detail.dataBind();
			
			/*****************************Detalle pagos  prepare the data*****************************/

            var dataFields_detail_pago = [
                        { name: 'fecha' },
                        { name: 'medio_pago' },
                        { name: 'cheque' },
						{ name: 'nombre_banco' },
                        { name: 'monto' },
                        { name: 'id_compra' }
                    ];
				
            var source_detail_pago =
            {
                datafields: dataFields_detail_pago,
                localdata: localdata_detail_all_pago
			};
            var dataAdapter_detail_pago = new $.jqx.dataAdapter(source_detail_pago);
            dataAdapter_detail_pago.dataBind();			
			/*****************************INICIO EVENTO CLICK MASTER**********************************/
            $("#comprasGrid").on('rowselect', function (event) {
                var id_compra = event.args.row.id_compra;
				
				/////////////LLENADO DEL DETALLE DE COMPRAS
                var records = new Array();
                var length = dataAdapter_detail.records.length;
                for (var i = 0; i < length; i++) {
                    var record = dataAdapter_detail.records[i];
                    if (record.id_compra == id_compra) {
                        records[records.length] = record;
                    }
                }
                var dataSource_aux = {
                    datafields: dataFields_detail,
                    localdata: records
                }
                var adapter_aux = new $.jqx.dataAdapter(dataSource_aux);
        
                // update data source.
                $("#detalleComprasGrid").jqxGrid({ source: adapter_aux });

				/////////////LLENADO DEL DETALLE DE PAGOS
                var records_pag = new Array();
                var length_pag = dataAdapter_detail_pago.records.length;
                for (var i = 0; i < length_pag; i++) {
                    var record = dataAdapter_detail_pago.records[i];
                    if (record.id_compra == id_compra) {
                        records_pag[records_pag.length] = record;
                    }
                }
                var dataSource_aux_pag = {
                    datafields: dataFields_detail_pago,
                    localdata: records_pag
                }
                var adapter_aux_pag = new $.jqx.dataAdapter(dataSource_aux_pag);
        
                // update data source.
                $("#detallePagosGrid").jqxGrid({ source: adapter_aux_pag });				
            });
			/*****************************FIN EVENTO CLICK MASTER*************************************/
			
			/*******CREACION GRILLA DETALLE COMPRAS*****************/
            $("#detalleComprasGrid").jqxGrid(
            {
                width: 480,
                height: 150,
                keyboardnavigation: false,				
                columns: [
                    { text: 'Articulo', datafield: 'descripcion_articulo', width: 150 },
                    { text: 'Cantidad', datafield: 'cantidad', width: 110 },
                    { text: 'Precio Unitario', datafield: 'precio_unitario', width: 110 },
                    { text: 'Subtotal', datafield: 'subtotal', width: 110 }
                ]
            });
			/*******CREACION GRILLA DETALLE PAGOS*******************/
            $("#detallePagosGrid").jqxGrid(
            {
                width: 650,
                height: 150,
                keyboardnavigation: false,				
                columns: [
                    { text: 'Fecha', datafield: 'fecha', cellsformat: 'd',width: 110 },
                    { text: 'Medio', datafield: 'medio_pago', width: 110 },
                    { text: 'Banco', datafield: 'nombre_banco', width: 150 },
					{ text: 'Cheque', datafield: 'cheque', width: 150 },
                    { text: 'Monto', datafield: 'monto', width: 110 }					
                ]
            });		
			/*******************************************************/			
            $("#comprasGrid").jqxGrid('selectrow', 0);			
	 
 }
 $(document).ready(function () {
            // prepare the data
			var localdata_master =  [];
			var localdata_detail_all = [];	
			var localdata_detail_all_pago = [];			
			$.get("query_compras_rep_JSON.asp", 
				function(data) {
									localdata_master = $.parseJSON(data)[0].data_master;
									localdata_detail_all = $.parseJSON(data)[1].data_detail;
									localdata_detail_all_pago = $.parseJSON(data)[2].data_detail_pagos;
									inicalizar(	localdata_master
												,localdata_detail_all
												,localdata_detail_all_pago
												);
							});									
        });
    
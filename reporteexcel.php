<?php
include 'conexion.php';
conectarse();
error_reporting(0);
$consulta="select get_mes(fecha_prefactura) as periodo, PR.provincia, C.ciudad, PA.parroquia, 
CL.razon_social, I.direccion_instalacion, 
REPLACE(substr(CL.telefono, 1, case when strpos(CL.telefono, '/')>0 then strpos(CL.telefono, '/')-1 else 10 end),'-',' ') as telefono,
'1','CNT','MEDIO INALAMBRICO',
PS.burst_limit as up_link, PS.burst_limit as down_link, replace(replace(replace(PS.solo_plan, 'ESPECIAL',' '),'SMALL','CIBERCAFE'),'NOCTURNO','RESIDENCIAL'), 
(PS.comparticion::int ||' : 1')  
from (((((tbl_instalacion as I inner join tbl_prefactura as P on P.id_instalacion=I.id_instalacion) 
inner join vta_provincia as PR on PR.id_provincia=I.id_provincia)
inner join vta_ciudad as C on C.id_ciudad=I.id_ciudad)
inner join vta_parroquia as PA on PA.id_parroquia=I.id_parroquia) 
inner join tbl_cliente as CL on I.id_cliente=CL.id_cliente) 
inner join vta_plan_servicio as PS on PS.id_plan_servicio=I.id_plan_actual 
where fecha_prefactura between '01-10-2014' and '31-10-2014' 
and P.fecha_emision is not null 
and PA.parroquia not in('CANCHAGUANO','EL CHAMIZO','MONTUFAR','GARCIA MORENO','LA MERCED','MIRA','SAN LUIS','SANTO DOMINGO','COTACACHI','6 DE JULIO','INTAG','SAN JOSE DE CHALTURA',
'ALPACHACA','AZAYA','EJIDO DE IBARRA','EL MILAGRO','EL TEJAR','LA FLORIDA','MILAGRO','SAN JOSE DE CHALTURA','SAN PEDRO','TANGUARIN','GUAICO PUNGO','ILUMAN BAJO','OTAVALO','SAGRARIO',
'SAN JUAN','SAN ROQUE','SAN VICENTE FERRER','SAN JOSE DE MINAS','ALADO DE LA CARPINTERIA','AMAZONAS','ATRAS DE LA ACADEMIA LA JOYA','BARRIO 15 DE ENERO','BARRIO 1 DE MAYO','BARRIO 25 DE DICIEMBRE',
'BARRIO 5 JUNIO','BARRIO 9 DE OCTUBRE','BARRIO AMAZONAS','BARRIO CENTRAL','BARRIO EL PARAISO','BARRIO JUMANDY','BARRIO LA ALBORADA','BARRIO LA LIBERTAD','BARRIO LIBERTAD',
'BARRIO LUZ DE AMERICA','BARRIO MACHALA','BARRIO MAGDALENA','BARRIO OSCAR ROMERO','BARRIO PRIMERO DE MAYO','BARRIO SANTA RITA','BARRIO SOL DE ORIENTE','COMUNIDAD 25 DE DICIEMBRE',
'COMUNIDAD LA PÁRKER','EL PORVENIR','ENTRADA A LOS ANGELES','HUAMAYACU','LA CAROLINA','LA LIBERTAD','LA MAGADALENA','LA PARKER','LA PIMAMPIRO','LOS LAURELES','MACHALA','SACHA','SACHAS',
'VALLE HERMOSO','AV. FERNANDO NOA Y JOSE LEIVA','AV. RAFAEL ANDRADE CHACON','BARRIO 5 AGOSTO','RAFAEL ANDRADE','BARRIO LAS AMERICAS','COMUNIDAD UNION MILAGREÑA','SAN SABASTIAN DEL COCA',
'CAJAS','CARRERA','CHAUPIESTANCIA','GRANOBLES','GUACHALA','LA JOSEFINA','MINAS','MIRAFLORES DEL YASNAN','MOYURCO','PEDRO MONCAYO','PINGULMI','SAN NICOLAS','SANTA ROSA DE PINGULMI',
'CAJAS JURIDICA','CANANVALLE','PEDRO MONCAYO','TOMALON','QUITO','BARRIO CENTRAL','BARRIO ESPEJO','LA UNION','LOS RIOS','ALADO DEL BANCO DEL FOMENTO','COCHA SECA','EL PLAYON',
'SANTA BARABARA','SAN ELOY','BARRIO SANTA ROSA','CAROLINA','SAN LUIS DE GUACHALA','CUBINCHE','CHIMBACALLE','PLOTARIA VARGAS JOSE RIVAS','BARRIO LA FLORIDA','SAN SEBASTIAN DEL COCA',
'BARRIO ELOY ALFARO')
and C.ciudad not in ('BOLIVAR')
order by (fecha_prefactura), PR.provincia, C.ciudad, PA.parroquia, CL.razon_social"; 

$resultado=pg_query($consulta) or die (pg_last_error());

	if(pg_num_rows($resultado)>0){

		if (PHP_SAPI == 'cli')
			die('Este archivo solo se puede ver desde un navegador web');

		/** Se agrega la libreria PHPExcel */
		require_once 'lib/PHPExcel/PHPExcel.php';

		// Se crea el objeto PHPExcel
		$objPHPExcel = new PHPExcel();

		// Se asignan las propiedades del libro
		$objPHPExcel->getProperties()->setCreator("Romero Giovanni-0995988018") //Autor
							 ->setLastModifiedBy("Saitel") //Ultimo usuario que lo modificó
							 ->setTitle("1.- Formato_subir_LineasDedi")
							 ->setSubject("Reporte Sietel")
							 ->setDescription("Reporte Lineas Dedicadas")
							 ->setKeywords("Lineas Dedicadas Saitel")
							 ->setCategory("Reportes Saitel a Sietel");

		$tituloReporte = "REPORTE DE SERVICIO CUENTAS ALTA VELOCIDAD";
		$subtituloReporte = "ACCESO NO CONMUTADO - LÍNEAS DEDICADAS";
		$subsubtituloReporte = "MES";
		$titulosColumnas = array('Provincia', 'Cantón', 'Parroquia', 'Nombre del Usuario', 'Dirección', 'Teléfono', 
			'Número estimado de usuarios por cuenta', 'Empresa proveedora del canal (Portador)', 
			'Tipo de enlace: Cobre, Cable Coaxial, Fibra Óptica, Medio Inalámbrico', 'Ancho de banda Up Link (Kbps)', 
			'Ancho de banda Down Link (Kbps)', 'Tipo de Cliente (Residencial, Corporativo, Cibercafé)', 'Nivel de Compartición');
		
		$objPHPExcel->setActiveSheetIndex(0)
        		    ->mergeCells('A1:N1')
        		    ->mergeCells('B2:N2')
        		    ->mergeCells('A2:A3');
						
		// Se agregan los titulos del reporte
		$objPHPExcel->setActiveSheetIndex(0)
					->setCellValue('A1',$tituloReporte)
					->setCellValue('B2',$subtituloReporte)
					->setCellValue('A2',$subsubtituloReporte)
        		    ->setCellValue('B3',  $titulosColumnas[0])
		            ->setCellValue('C3',  $titulosColumnas[1])
        		    ->setCellValue('D3',  $titulosColumnas[2])
            		->setCellValue('E3',  $titulosColumnas[3])
            		->setCellValue('F3',  $titulosColumnas[4])
            		->setCellValue('G3',  $titulosColumnas[5])
            		->setCellValue('H3',  $titulosColumnas[6])
            		->setCellValue('I3',  $titulosColumnas[7])
            		->setCellValue('J3',  $titulosColumnas[8])
            		->setCellValue('K3',  $titulosColumnas[9])
            		->setCellValue('L3',  $titulosColumnas[10])
            		->setCellValue('M3',  $titulosColumnas[11])
            		->setCellValue('N3',  $titulosColumnas[12]);
		
		//Se agregan los datos de los alumnos
		$i = 4;
		while ($fila =pg_fetch_array($resultado)) {
			$objPHPExcel->setActiveSheetIndex(0)
        		    ->setCellValue('A'.$i,  $fila[0])
		            ->setCellValue('B'.$i,  $fila[1])
        		    ->setCellValue('C'.$i,  $fila[2])
            		->setCellValue('D'.$i, 	$fila[3])
            		->setCellValue('E'.$i,  $fila[4])
            		->setCellValue('F'.$i,  $fila[5])
            		->setCellValue('G'.$i,  $fila[6])
            		->setCellValue('H'.$i,  $fila[7])
            		->setCellValue('I'.$i,  $fila[8])
            		->setCellValue('J'.$i,  $fila[9])
            		->setCellValue('K'.$i,  $fila[10])
            		->setCellValue('L'.$i,  $fila[11])
            		->setCellValue('M'.$i,  $fila[12])
            		->setCellValue('N'.$i,  $fila[13]);
					$i++;
		}
		
		$estiloTituloReporte = array(
        	'font' => array(
	        	'name'      => 'Calibri',
    	        'bold'      => false,
        	    'italic'    => false,
                'strike'    => false,
               	'size' =>12
	            //'color'     => array('rgb' => '96BEE6')
            ),
	        'fill' => array(
				'type'	=> PHPExcel_Style_Fill::FILL_SOLID,
				'color'	=> array('argb' => '96BEE6')
			),
            'borders' => array(
               	'allborders' => array(
                	'style' => PHPExcel_Style_Border::BORDER_THIN                    
               	)
            ), 
            'alignment' =>  array(
        			'horizontal' => PHPExcel_Style_Alignment::HORIZONTAL_CENTER,
        			'vertical'   => PHPExcel_Style_Alignment::VERTICAL_CENTER
        			//'wrap'          => TRUE
    		)
        );

		$estiloTituloColumnas = array(
            'font' => array(
                'name'      => 'Arial',
                'bold'      => true,                          
                'color'     => array(
                    'rgb' => 'FFFFFF'
                )
            ),
            'fill' 	=> array(
				'type'		=> PHPExcel_Style_Fill::FILL_GRADIENT_LINEAR,
				'rotation'   => 90,
        		'startcolor' => array(
            		'rgb' => 'c47cf2'
        		),
        		'endcolor'   => array(
            		'argb' => 'FF431a5d'
        		)
			),
            'borders' => array(
            	'top'     => array(
                    'style' => PHPExcel_Style_Border::BORDER_MEDIUM ,
                    'color' => array(
                        'rgb' => '143860'
                    )
                ),
                'bottom'     => array(
                    'style' => PHPExcel_Style_Border::BORDER_MEDIUM ,
                    'color' => array(
                        'rgb' => '143860'
                    )
                )
            ),
			'alignment' =>  array(
        			'horizontal' => PHPExcel_Style_Alignment::HORIZONTAL_CENTER,
        			'vertical'   => PHPExcel_Style_Alignment::VERTICAL_CENTER,
        			'wrap'          => TRUE
    		));
			
		$estiloInformacion = new PHPExcel_Style();
		$estiloInformacion->applyFromArray(
			array(
           		'font' => array(
               	'name'      => 'Arial',               
               	'color'     => array(
                   	'rgb' => '000000'
               	)
           	),
           	'fill' 	=> array(
				'type'		=> PHPExcel_Style_Fill::FILL_SOLID,
				'color'		=> array('argb' => 'FFd9b7f4')
			),
           	'borders' => array(
               	'left'     => array(
                   	'style' => PHPExcel_Style_Border::BORDER_THIN ,
	                'color' => array(
    	            	'rgb' => '3a2a47'
                   	)
               	)             
           	)
        ));
		 
		$objPHPExcel->getActiveSheet()->getStyle('A1:N1')->applyFromArray($estiloTituloReporte);
		$objPHPExcel->getActiveSheet()->getStyle('A2:N2')->applyFromArray($estiloTituloReporte);
		$objPHPExcel->getActiveSheet()->getStyle('A3:N3')->applyFromArray($estiloTituloReporte);
		//$objPHPExcel->getActiveSheet()->getStyle('A3:D3')->applyFromArray($estiloTituloColumnas);		
		//$objPHPExcel->getActiveSheet()->setSharedStyle($estiloInformacion, "A4:D".($i-1));
        
		for($i = 'A'; $i <= 'N'; $i++){
			$objPHPExcel->setActiveSheetIndex(0)			
				->getColumnDimension($i)->setAutoSize(TRUE);
		}
		
		// Se asigna el nombre a la hoja
		$objPHPExcel->getActiveSheet()->setTitle('Hoja1');

		// Se activa la hoja para que sea la que se muestre cuando el archivo se abre
		$objPHPExcel->setActiveSheetIndex(0);
		// Inmovilizar paneles 
		//$objPHPExcel->getActiveSheet(0)->freezePane('A4');
		$objPHPExcel->getActiveSheet(0)->freezePaneByColumnAndRow(0,4);

		// Se manda el archivo al navegador web, con el nombre que se indica (Excel2007)
		header('Content-Type: application/vnd.openxmlformats-officedocument.spreadsheetml.sheet');
		header('Content-Disposition: attachment;filename="1.- Formato_subir_LineasDedi.xlsx"');
		header('Cache-Control: max-age=0');

		$objWriter = PHPExcel_IOFactory::createWriter($objPHPExcel, 'Excel2007');
		$objWriter->save('php://output');
		exit;
		
	}
	else{
		print_r('No hay resultados para mostrar');
	}
?>
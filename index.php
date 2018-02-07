<?php

require_once __DIR__ . '/lib/PHPReporter/PHPReporter.php';


$csv = array(
	'headers' => array('Num. Empleado', 'Genero', 'Nombre'),
	'cols' => array('emp_no','gender','first_name'),
	'rows' => array(
		array(
			'emp_no' => 10002,
	        'first_name' => 'Bezalel',
	        'last_name' => 'Simmel',
	        'gender' => 'F'
		),
		array(
			'emp_no' => 10003,
		    'first_name' => 'Parto',
		    'last_name' => 'Bamford',
		    'gender' => 'M'
		)
	)
);

$data = array(
	array(
		'range' => 'A1',
		'data' => 'Esto es un super texto de pruebas para cekda con autoajuste',
		'props' => array(
			'text' => array(
				'font' => 'Calibri',
				'size' => '12',
				'bold' => true,
				'italic' => true,
				'cursive' => false,
				'color' => 'red',
				'colorHex' => 'A2186D',
				'underlined' => false
			),
			'cell' => array(
				'autoSize' => false,
				'wrapText' => true
			)
			
		)
	),
	array(
		'range' => 'B1',
		'data' => 'Celda sin formato'
	),
	array(
		'range' => 'A2:D2',
		'data' => 'Titulo del reporte',
		'props' => array(
			'text' => array(
				'red' => 'Black',
				'size' => '24'
			),
			'cell' => array(
				'merge' => true
			)
		)
	),
	array(
		'range' => 'A$',
		'data' => 'Celda variable'
	),
	array(
		'range' => 'A4:D4',
		'data' => array('Col1', 'Col2', 'Col3', 'Col4')
	),
	 array(
	 	'range' => 'A$:D$',
		'data' => 'Rango variable y merge',
		'props' => array(
			'text' => array(
				'red' => 'Black',
				'size' => '24'
			),
			'cell' => array(
				'merge' => true
			)
		)
	),
	array(
		'range' => 'A$:D$',
		'data' => array('Val1', 'Val2', 'Val3', 'Val4', 'Val5'),
		'props' => array(
			'text' => array(
				'red' => 'Black',
				'size' => '14'
			)
		)
	),
	array(
		'range' => 'A12:$',
		'data' => array(
			array('1','2','3','4'),
			array('1','2','3','4'),
			array('1','2','3','4'),
			array('1','2','3','4'),
			array('1','2','3','4')
		)
	),
	array(
		'range' => 'A$:$',
		'data' => array(
			array('Dat1','Dat2','Dat3','Dat4'),
			array('Dat1','Dat2','Dat3','Dat4')
		)
	)
);

$params = array(
	'nombre' => 'MiReporte',
	'titulo' => 'Mi Reporte´',
	'autor' => 'A. Venegas',
	'asunto' => 'A. Venegas',
	'descripcion' => 'Documento de Andrés Venegas',
	'palabrasclave' => 'Andres, Venegas, Documento',
	'categoria' => 'Desconocida'
);

// echo '<pre>';
// print_r($data);
// echo '</pre>';
// die();

$obj = new PHPReporter\PHPReporter();
$obj->displayReport('xlsx', $data, $params);
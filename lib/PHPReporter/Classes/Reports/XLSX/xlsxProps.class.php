<?php

namespace PHPReporter\Reports\XLSX;

/**
 * XLSX Properties
 * 
 * @author Andrés Venegas R.
 * @package xlsxReporter
 * 
 */
class xlsxProps {
	/**
	 * propiedades
	 * 
	 * @var array con las propiedades básicas de una o un rango de celdas.
	 */
	protected static $propiedades = array(
		'text' => array(
			'font' => 'Arial', // Tipos de letra estándar de windows
			'size' => 11, // Tamaño en puntos
			'bold' => false, // Verdadero o falso para activar la propiedad
			'italic' => false, // Verdadero o falso para activar la propiedad
			'underlined' => false, // Verdadero o falso para activar la propiedad
			'color' => 'black', // BLACK,WHITE,RED,DARKRED,BLUE,DARKBLUE,GREEN,DARKGREEN,YELLOW,DARKYELLOW
			'colorHex' => ''
		),
		'cell' => array(
			'autoSize' => false,
			'wrapText' => false,
			'merge' => false
		)
	);
}
<?php

include 'workbook.php';
include 'styles.php';

$Workbook = new Workbook();
$Styles = new Styles($Workbook);

$Worksheet = $Workbook->addWorksheet('Sheet1');
$Table = $Worksheet->addTable();
$Row = $Table->addRow(array('ID', 'Firstname', 'Lastname'));

$Table->addRows
(
	array
	(		
		array(1, 'Christoff', 'Truter'),
		array(2, 'Eugene', 'Stander')
	)
);

$Style = $Styles->setStyle('A', $Row);
$Style->SetFont('Times', '16', 'Red', true);

$Workbook->Output();


?>
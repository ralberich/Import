<html>

<?php
	
	include("connect.php");
	include("fonctions.php");
	//include("inc/variable.inc.php");
	
	//	include("inc/header.inc.php");

	

	
?>
<head>
<meta http-equiv="Content-Type" content="text/html; charset=utf8" />

<title>Import Incidents</title>

</head>
<div id="main">

	<?php
	require_once 'Classes/PHPExcel/IOFactory.php';
	// Chargement du fichier Excel
	$objPHPExcel = PHPExcel_IOFactory::load("EXPORT_IM-2.xls");
 
/**
* récupération de la première feuille du fichier Excel
* @var PHPExcel_Worksheet $sheet
*/
$sheet = $objPHPExcel->getSheet(0);
$entree = array();


//détermination des limites du tableau excel
$highestRow = $sheet->getHighestRow(); 
$highestColumn = $sheet->getHighestColumn();
//transformation du numéro de colonne (lettré) en chiffres : 
$highestColumnIndex = PHPExcel_Cell::columnIndexFromString($highestColumn); 


//parcours de toutes les lignes en ignorant la première
for ($row = 2; $row <= $highestRow; ++$row) {

	//récupération du numéro d'IM du fichier excel
	$numIMxl = ($sheet->getCellByColumnAndRow(0, $row)->getValue());
	//récupération de la date de maj de l'IM sur l'excel en convertissant en format timestamp
	$datemajxl = PHPExcel_Shared_Date::ExcelToPHP($sheet->getCellByColumnAndRow(2, $row)->getValue());
	
	//requete en base a partir du code IM pour identifier si celui ci existe déjà, date de maj et ID
	$req = "SELECT count(ID) as NB, max(ID) as ID, max(date_maj) as date_maj FROM incidents WHERE numero = '".$numIMxl."';";
	$res = $bdd->query($req);
	$IMexistant = $res->fetch();
	echo $numIMxl."  existant : ".$IMexistant['NB']."</br>";
	
	
	//Si l'IM existe déjà
	if($IMexistant['NB'] >= 1){
		$lastID=$IMexistant['ID'];
		
		//conversion des dates de maj au format datetime
		$dateIM = new DateTime($IMexistant['date_maj']);
		$dateXL = new DateTime();
		$dateXL->setTimestamp($datemajxl);
		
		//si la date de Maj du fichier excel est supérieur à la date de maj de l'IM en base, on met à jour la base.
		if ($dateIM->format('YmdHis') < $dateXL->format('YmdHis')){
			$entree = extract_ligne($sheet,$row,$highestColumnIndex);
			update_IM($bdd,$lastID,$entree);
			create_MAJ($lastID,$entree,$dateIM,$bdd);
		}
	}
	else {
		$entree = extract_ligne($sheet,$row,$highestColumnIndex);
		$lastID=create_IM($bdd,$entree);
		$dateblank = new DateTime('01/01/1980');
		create_MAJ($lastID,$entree,$dateblank,$bdd);
	}
}
	
	?>

</div>
</body>

</html>

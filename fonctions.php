<?php
function extract_ligne($sheet,$row,$highestColumnIndex)
{ //fonction permettant d'extraire une ligne d'un fichier excel, pour une ligne donnée, étant fourni la colonne maxi au format numérique
	$entree = array();
	// On boucle sur les cellule de la ligne
	for ($col = 0; $col <= $highestColumnIndex; ++$col) {
		//on affecte la cellule
		$cell = ($sheet->getCellByColumnAndRow($col, $row));
		
		//Pour chaque colonne on affecte au tableau $entree
		switch($cell->GetColumn()){
			case "A":
				$entree['numero']=$cell->Getvalue();
			break;
			case "B":
				$entree['date_ouverture']=PHPExcel_Shared_Date::ExcelToPHP($cell->getvalue());
			break;
			case "C":
				$entree['date_maj']=PHPExcel_Shared_Date::ExcelToPHP($cell->getvalue());
			break;
			case "D":
				$testdate = $cell->Getvalue();
				if($testdate!=""){
					$entree['date_cloture']=PHPExcel_Shared_Date::ExcelToPHP($cell->getvalue());
				}
				else{
					$entree['date_cloture']="";
				}
				
			break;
			case "E":
				$entree['titre']=$cell->Getvalue();
			break;
			case "F":
				$entree['description']=$cell->Getvalue();
			break;
			case "G":
				$entree['action_maj']=$cell->Getvalue();
			break;
			case "H":
				$entree['solution']=$cell->Getvalue();
			break;
			case "I":
				$entree['etat']=$cell->Getvalue();
			break;
			case "J":
				$entree['ci']=$cell->Getvalue();
			break;
			case "K":
				$entree['user_creation']=$cell->Getvalue();
			break;
			case "L":
				$entree['user_maj']=$cell->Getvalue();
			break;
			case "M":
				$entree['user_cloture']=$cell->Getvalue();
			break;
			case "N":
				$entree['groupe_createur']=$cell->Getvalue();
			break;
			case "O":
				$entree['groupe_affectation']=$cell->Getvalue();
			break;
			case "P":
				$entree['groupe_cloture']=$cell->Getvalue();
			break;
			case "U":
				$entree['priorite']=$cell->Getvalue();
			break;
			case "V":
				$entree['code_cloture']=$cell->Getvalue();
			break;
			case "W":
				$entree['fournisseur']=$cell->Getvalue();
			break;
			case "X":
				$entree['ref']=$cell->Getvalue();
			break;
			case "Z":
				$entree['element_ci']=$cell->Getvalue();
			break;
			case "AA":
				$entree['user_beneficiaire']=$cell->Getvalue();
			break;
			case "AB":
				$entree['candidat_bdd']=$cell->Getvalue();
			break;
			case "AD":
				$entree['user_email']=$cell->Getvalue();
			break;
			case "AE":
				$entree['user_tel_fixe']=$cell->Getvalue();
			break;
			case "AF":
				$entree['user_tel_portable']=$cell->Getvalue();
			break;
		}
		
	}
	//on renvoie le tableau contenant les résultats
	return $entree;
	
}

function controle_user($user,$bdd){

 /*	Si NNI présent dans user, controle sur le NNI OU sur la combinaison Nom Prenom
		Si le NNI est vide en base mais l'utilisateur existe, on met à jour le NNI en base
	Si uniquement nom prenom, controle sur nom / prenom
 
	Dans tous les cas, si l'user est trouvé, on renvoie l'ID correspondant
	
	Si l'user n'est pas retrouvé, on le crée et on revoie l'ID crée. */
	if($user ==""){
		return "";
	}

	//recherche en base pour voir si il existe la personne
	$req = "SELECT count(ID) as NB , max(ID) as ID, max(nni) as NNI FROM users ";
	
	//si le NNI est présent ds le fichier excel, on recherche une correspondance avec le NNI ou la combinaison nom/prenom
	if(stripos($user,"(")) {
		$nni = substr($user,stripos($user,"(")+1,(stripos($user,")")-stripos($user,"("))-1);
		$nom = substr($user,0,stripos($user,","));
		$prenom = substr($user,stripos($user,",")+2,(stripos($user,"(")-stripos($user,","))-3);
		$req .= "WHERE NNI = '" .$nni. "' OR (nom = '" .$nom. "' AND prenom = '".$prenom."');";
	}
	
	//si le NNI n'est pas présent ds le fichier excel, on recherche une correspondance uniquement avec la combinaison nom/prenom
	else{
		$nni = "";
		$prenom = substr($user,0,stripos($user," "));
		$nom = substr($user,stripos($user," ")+1);
		
		$req .= "WHERE nom = '" .$nom. "' AND prenom = '".$prenom."';";
	}
	//requete en base
	$res = $bdd->query($req);
	$userexistant = $res->fetch();
	
	//si le NNI existe en base
	if($userexistant['NB'] >= 1){
		
		//si l'entrée en base de l'utilisateur trouvé n'a pas de NNI entré et que le fichier excel en a une, on le met à jour.
		if($userexistant['NNI']=="" and $nni!=""){
			$req_nni="UPDATE users SET nni = '".$nni."' WHERE ID = '".$userexistant['ID']."';";
			$bdd->exec($req_nni);
		}
		//on renvoie l'ID existant
		return $userexistant['ID'];
	}
	else{
		//si l'utilisateur n'existe pas, on l'injecte.
		$req = "INSERT INTO users (";
		$req .= "nni,nom,prenom) ";
		$req .= "VALUES ('".$nni."','".$nom."','".$prenom."');";
		$req;
		$bdd->exec($req);
		//on renvoie le numero de l'ID de l'enregistrement crée.
		return $bdd->lastInsertId();
	}
}
function controle_groupe($groupe,$bdd){
	if($groupe=="") {return "";}
	$req = "SELECT count(ID) as NB , max(ID) as ID, max(libelle) as libelle FROM groupes ";
	$req .= "WHERE libelle = '".$groupe."';";
	$res = $bdd->query($req);
	$groupeexistant = $res->fetch();
	
	if($groupeexistant['NB'] >= 1){
		return $groupeexistant['ID'];
	}
	else {
		$req = "INSERT INTO groupes (";
		$req .= "libelle)";
		$req .= "VALUES('".$groupe."')";
		
		$bdd->exec($req);
		return $bdd->lastInsertId();
	}
}

function controle_type_maj($type,$bdd){
	if($type=="") {return "";}
	$req = "SELECT count(ID) as NB , max(ID) as ID, max(libelle) as libelle FROM types_maj ";
	$req .= "WHERE libelle = '".addslashes($type)."';";
	
	$res = $bdd->query($req);
	$typeexistant = $res->fetch();
	
	if($typeexistant['NB'] >= 1){
		return $typeexistant['ID'];
	}
	else {
		$req = "INSERT INTO types_maj (";
		$req .= "libelle)";
		$req .= "VALUES('".addslashes($type)."')";
		
		$bdd->exec($req);
		return $bdd->lastInsertId();
	}
}

function update_IM($bdd,$ID,$entree){
//fonction permettant de mettre à jour un IM déjà existant, à partir de l'ID en base et d'un tableau contenant les valeurs à insérer

//pour chaque user, controle de l'existence et création si nécessaire
$id_user_creation = controle_user($entree['user_creation'],$bdd);
$id_user_maj = controle_user($entree['user_maj'],$bdd);
$id_user_cloture = controle_user($entree['user_cloture'],$bdd);
$id_user_beneficiaire = controle_user($entree['user_beneficiaire'],$bdd);

//pour chaque groupe, controle de l'existence et création si nécessaire
$id_groupe_createur = controle_groupe($entree['groupe_createur'],$bdd);
$id_groupe_affectation = controle_groupe($entree['groupe_affectation'],$bdd);
$id_groupe_cloture = controle_groupe($entree['groupe_cloture'],$bdd);

$req = "UPDATE incidents SET ";
$req .= "date_ouverture = '".date("Y-m-d H:i:s",$entree['date_ouverture'])."' , ";
$req .= "date_maj = '".date("Y-m-d H:i:s",$entree['date_maj'])."' , ";
if($entree['date_cloture']!= ""){
	$req .= "date_cloture = '".date("Y-m-d H:i:s",$entree['date_cloture'])."' , ";
}
$req .= "titre = '".utf8_encode(addslashes($entree['titre']))."' , ";
$req .= "description = '".addslashes($entree['description'])."' , ";
$req .= "solution = '".addslashes($entree['solution'])."' , ";
$req .= "etat = '".$entree['etat']."' , ";
$req .= "ci = '".$entree['ci']."' , ";
$req .= "id_user_creation = '".$id_user_creation."' , ";
$req .= "id_user_cloture = '".$id_user_cloture."' , ";
//manque id premier user sn2
$req .= "id_groupe_createur = '".$id_groupe_createur."' , ";
$req .= "id_groupe_affectation = '".$id_groupe_affectation."' , ";
$req .= "id_groupe_cloture = '".$id_groupe_cloture."' , ";
$req .= "priorite = '".$entree['priorite']."' , ";
$req .= "code_cloture = '".$entree['code_cloture']."' , ";
$req .= "fournisseur = '".$entree['fournisseur']."' , ";
$req .= "ref = '".$entree['ref']."' ";

$req .= "WHERE ID = ".$ID;
//echo $req;
//$bdd->exec($req);

}


function create_IM($bdd,$entree){
//fonction permettant de créer un IM n'existant pas encore en base a partir d'un tableau contenant les valeurs à insérer

//pour chaque user, controle de l'existence et création si nécessaire
$id_user_creation = controle_user($entree['user_creation'],$bdd);
$id_user_maj = controle_user($entree['user_maj'],$bdd);
$id_user_cloture = controle_user($entree['user_cloture'],$bdd);
$id_user_beneficiaire = controle_user($entree['user_beneficiaire'],$bdd);

//pour chaque groupe, controle de l'existence et création si nécessaire
$id_groupe_createur = controle_groupe($entree['groupe_createur'],$bdd);
$id_groupe_affectation = controle_groupe($entree['groupe_affectation'],$bdd);
$id_groupe_cloture = controle_groupe($entree['groupe_cloture'],$bdd);

$req = "INSERT INTO incidents (";
$req .= "numero, date_ouverture, date_maj, date_cloture, titre, description, solution, etat, ci ";
$req .= ", id_user_creation, id_user_cloture, id_groupe_createur, id_groupe_affectation, id_groupe_cloture, priorite, code_cloture, fournisseur, ref";
$req .= " ) VALUES (";
$req .= "'".$entree['numero']."' , ";
$req .= "'".date("Y-m-d H:i:s",$entree['date_ouverture'])."' , ";
$req .= "'".date("Y-m-d H:i:s",$entree['date_maj'])."' , ";
if($entree['date_cloture']!= ""){
	$req.= "'".date("Y-m-d H:i:s",$entree['date_cloture'])."' , ";
}
else {
	$req .= " '' , ";
}
$req .= "'".utf8_encode(addslashes($entree['titre']))."' , ";
$req .= "'".addslashes($entree['description'])."' , ";
$req .= "'".addslashes($entree['solution'])."' , ";
$req .= "'".$entree['etat']."' , ";
$req .= "'".$entree['ci']."' , ";
$req .= "'".$id_user_creation."' , ";
$req .= "'".$id_user_cloture."' , ";
//manque id premier user sn2
$req .= "'".$id_groupe_createur."' , ";
$req .= "'".$id_groupe_affectation."' , ";
$req .= "'".$id_groupe_cloture."' , ";
$req .= "'".$entree['priorite']."' , ";
$req .= "'".$entree['code_cloture']."' , ";
$req .= "'".$entree['fournisseur']."' , ";
$req .= "'".$entree['ref']."' ";
$req .= ")";
//echo $req;
$bdd->exec($req);
return $bdd->lastInsertId();

}

function create_MAJ($ID,$entree,$datetest,$bdd){
	$majs=array();
	$lignemaj=array();
	$majs=explode(" jour le ",$entree['action_maj']);
	//pour chacune des maj
	for($i=1;$i<sizeof($majs);$i++){
	
	//on controle que la date soit plus récente si datemaj n'est pas à 0 (cas de création)
		$datemaj = DateTime::createFromFormat("d/m/y G:i:s",substr($majs[$i],0,17));
		if(($datemaj->format('YmdHis') > $datetest->format('YmdHis'))) {
			$lignemaj['date']=$datemaj;
			
			//extraction du type de mise à jours
			$debut = 20;
			$long = stripos($majs[$i],"\n")-$debut;
			$lignemaj['type']=substr($majs[$i],$debut,$long);
			
			//extraction de l'utilisateur.
			$debut = stripos($majs[$i],"\n")+1;
			$long = stripos($majs[$i]," - ",stripos($majs[$i],"\n"))-$debut;
			$lignemaj['user']=substr($majs[$i],$debut,$long);
			
			//extraction du groupe de l'utilisateur
			$debut = $debut + $long + 3;
			$long = stripos($majs[$i]," :",$debut)-$debut;
			$lignemaj['groupe']=substr($majs[$i],$debut,$long);
			
			//extraction du contenu de la MAJ
			$debut = $debut + $long + 4;
			$long = stripos($majs[$i],"\n------")-$debut-1;
			$lignemaj['maj']=substr($majs[$i],$debut,$long);
			
			//test si la ligne du type contient une escalade.
			if (stripos($lignemaj['type'],"-->")){
				$debut = stripos($lignemaj['type'],"--> ")+4;
				$long = strlen($lignemaj['type'])-$debut -2 ;
				$lignemaj['groupe_dest'] = substr($lignemaj['type'],$debut,$long);
				$long = stripos($lignemaj['type']," de ");
				$lignemaj['type'] = substr($lignemaj['type'],0,$long);
				}
			else{
				$lignemaj['groupe_dest'] = $lignemaj['groupe'];
			}
			
			$id_user = controle_user($lignemaj['user'],$bdd);
			$id_groupe_origine = controle_groupe($lignemaj['groupe'],$bdd);
			$id_groupe_destination = controle_groupe($lignemaj['groupe_dest'],$bdd);
			$id_type_maj = controle_type_maj($lignemaj['type'],$bdd);
			$req = "INSERT INTO majs (";
			$req.= "id_incident, id_user, id_groupe_origine, id_groupe_destination";
			$req.= ", texte, date, id_type";
			$req.= ") VALUES (";
			$req.= "'".$ID."',";
			$req.= "'".$id_user."',";
			$req.= "'".$id_groupe_origine."',";
			$req.= "'".$id_groupe_destination."',";
			$req.= "'".addslashes($lignemaj['maj'])."',";
			$req.= "'".$lignemaj['date']->format("Y-m-d H:i:s")."',";
			$req.= "'".$id_type_maj."'";
			$req.= ");";
			//echo $req."</br>";
			$bdd->exec($req);
			//return $bdd->lastInsertId();
		}
	}
}

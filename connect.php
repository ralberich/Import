<?php
if (isset ($_GET["page"])){
$page = $_GET["page"];}
else {
$page = "projet";
}
/* Database config */
$db_host		= '127.0.0.1';
$db_user		= 'SN2';
$db_pass		= 'sn2sopra';
$db_database	= 'activite'; 

/* End config */
//echo "'mysql:host=".$db_host.";dbname=".$db_database."', '".$db_user."','".$db_pass."'";
try
{
$bdd = new PDO('mysql:host=127.0.0.1;dbname=activite', 'SN2','sn2sopra');
$bdd->query("SET NAMES 'utf8'");
}
catch (Exception $e)
{
die('Erreur : ' . $e->getMessage());
}


//$link = @mysql_connect($db_host,$db_user,$db_pass) or die('Unable to establish a DB connection');

//mysql_set_charset('utf8');
//mysql_select_db($db_database,$link);



?>
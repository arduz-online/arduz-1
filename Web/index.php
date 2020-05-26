<?php
	session_start();
	ob_start();
	
	
	$dbconn = mysql_connect("localhost","noicoder_root","fedegay");
	mysql_select_db("noicoder_sake",$dbconn);
	if ($_SESSION['loggedE']!="1"){
	$_SESSION['loggedE']="0";
	$_SESSION['passwd']="";
	}
	$actionArray = array(
		'm' => 						array('bin/index.php'),
		'ranking' => 				array('bin/rank.php'),
		'registrar-personaje' => 	array('bin/register.php'),
		'mi_cuenta' => 				array('bin/mi.php'),
		'descargar' => 				array('bin/do.php'),
		'login' => 					array('bin/login.php'),
		'panel' => 					array('bin/panel/panel.php'),
		'panel-clan' => 			array('bin/panel/clan.php'),
		'panel-crear-clan' => 		array('bin/panel/crearclan.php'),
		'panel-editar-clan' => 		array('bin/panel/clan.php'),
		'panel-ver-items' => 		array('bin/panel/misitems.php'),
		'panel-comprar-items' => 	array('bin/panel/compraritems.php'),
		'panel-elejir-items' =>		array('bin/panel/misitems.php'),
		'admin-panel' =>			array('bin/panel/misitems.php'),
		'equipo' =>			array('bin/staff.php'),
	);
	
	if ($_REQUEST['a'] =="salir")
	{
		session_destroy();
		header("Location: index.php");
	}
	if ($_REQUEST['a'] == "udr")// && $_REQUEST['b'] == "bbmanda" && $_REQUEST['c'] == "sabelo")
	{
		rank_update();
		echo mysql_error();
		$_REQUEST['a']="ranking";
	}
	
	foreach($_GET as $variable=>$valor){
	// Modifica las variables pasadas por URL
		$_GET[$variable] = @mysql_real_escape_string($valor);
	}
	foreach($_POST as $variable=>$valor){
	// Modifica las variables de formularios 
		$_POST[$variable] = @mysql_real_escape_string($valor);
	}
	foreach($_REQUEST as $variable=>$valor){
	// Modifica las variables de formularios y url
		$_REQUEST[$variable] = @mysql_real_escape_string($valor);
	}
	
	require_once 'bin/theme.php';
	if (!isset($_REQUEST['a']) || !isset($actionArray[$_REQUEST['a']]))
	{
	require_once 'bin/index.php';
	}else{
	require_once ($actionArray[$_REQUEST['a']][0]);
	}
	ob_end_flush();
	
	function bbcode($texto) {
		// aplica bbcode a la string $texto
		$texto = str_replace("\\n","<br/>",$texto); //Saltos de Linea
		$texto = str_replace("\n","<br/>",$texto); //Saltos de Linea
		$texto = str_replace("[b]","<b>",$texto); //Negrita
		$texto = str_replace("[/b]","</b>",$texto); //Negrita
		$texto = str_replace("[i]","<i>",$texto); //Italic
		$texto = str_replace("[/i]","</i>",$texto); //Italic
		$texto = str_replace("[u]","<u>",$texto); //Subrayado
		$texto = str_replace("[/u]","</u>",$texto); //Subrayado
		$texto = str_replace("[t]","<del>",$texto); //Tachado
		$texto = str_replace("[/t]","</del>",$texto); //Tachado
		$texto = str_replace("ñ","&ntilde;",$texto);
		$texto = str_replace("á","&aacute;",$texto);
		$texto = str_replace("é","&eacute;",$texto);
		$texto = str_replace("í","&iacute;",$texto);
		$texto = str_replace("ó","&oacute;",$texto);
		$texto = str_replace("ú","&uacute;",$texto);
		$texto = str_replace("Ñ","&Ntilde;",$texto);
		$texto = str_replace("Á","&Aacute;",$texto);
		$texto = str_replace("É","&Eacute;",$texto);
		$texto = str_replace("Í","&Iacute;",$texto);
		$texto = str_replace("Ó","&Oacute;",$texto);
		$texto = str_replace("Ú","&Uacute;",$texto);
		return $texto;
	}
	
	function email_is_valid($email)
	{
		if (substr_count($email, '@') != 1)
			return false;
		if ($email{0} == '@')
			return false;
		if (substr_count($email, '.') < 1)
			return false;
		if (strpos($email, '..') !== false)
			return false;
		$length = strlen($email);
		for ($i = 0; $i < $length; $i++) {
			$c = $email{$i};
			if ($c >= 'A' && $c <= 'Z')
				continue;
			if ($c >= 'a' && $c <= 'z')
				continue;
			if ($c >= '0' && $c <= '9')
				continue;
			if ($c == '@' || $c == '.' || $c == '_' || $c == '-')
				continue;
			return false;
		}
		$TLD = array ('COM',   'NET', 'ORG',   'MIL', 'EDU',   'GOV', 'BIZ',   'NAME', 'MOBI',  'INFO', 'AERO',  'JOBS', 'MUSEUM');
		$tld = strtoupper(substr($email, strrpos($email, '.') + 1));
		if (strlen($tld) != 2 && !in_array($tld, $TLD))
			return false;
		return true;
	}
	
	function u_online($minutos = 15){
		global $context;
		$tiem=$minutos * 60;
		$waa=mysql_fetch_array(mysql_query("SELECT COUNT(DISTINCT(id)) FROM `pjs` WHERE ultimologin>".(time()-$tiem)));
		$online=$waa[0];
		$context['onlineusers']=$online;
		return $online;
	}
	
	function u_online_t($minutos = 15){
		global $context;
		$result = mysql_query("SELECT * FROM `servers` ORDER BY `players` DESC");
						while ($row=mysql_fetch_array($result))
						{
							$online+=intval($row['players']);
						}
		return $online;
	}
	
	function rank_update()
	{
		$start = 1; 
		$query = mysql_query('SELECT * FROM `pjs` ORDER BY puntos DESC;'); 
	    while ($row = mysql_fetch_assoc($query)){ 
	       $playername_rank =  $row['nick']; 
	       $rank_old = $row['rank']; 
	       $start++; 
	       $query_rank = mysql_query("UPDATE `pjs` SET `rank_old`='".$rank_old."',`rank`='".$start."' WHERE `nick` = '".$playername_rank."'");       
	    }
		$start = 1; 
		$query = mysql_query('SELECT * FROM `pjs` ORDER BY frags DESC;'); 
	    while ($row = mysql_fetch_assoc($query)){ 
	       $playername_rank =  $row['nick']; 
	       $rank_old = $row['rank']; 
	       $start++; 
	       $query_rank = mysql_query("UPDATE `pjs` SET `rank_frags_old`='".$rank_old."',`rank_frags`='".$start."' WHERE `nick` = '".$playername_rank."'");       
	    }
		mysql_query("UPDATE `configuracion` SET `num` = '".$start."',`ultimoupd` = '".time()."' WHERE `cfg`='s' LIMIT 1"); 
	}

	function existeuser($user)
	{
		$sql = "SELECT `ID`,`nick` FROM `pjs` WHERE `nick`='".$user."' LIMIT 1;";
		$result=mysql_query($sql);
		$row=mysql_fetch_array($result);
		if (mysql_num_rows($result)>0)
		{
			return "1";
		} else {
			return "0";
		}
	}
	
	function infopj($datos)
	{
					$add="<br/>Ultima vez online: ";
					$tiempo = intval((time()-$datos['ultimologin'])/60);
					if ($tiempo < 61){
						  if ($tiempo > 2) {
						  	    $add .= "<b style='color:cyan;'>Hace ".$tiempo. " minutos.</b>";
						  } else {
						  	    $add .= "<b style='color:cyan;'>Hace unos instantes!.</b>";
						  }

						  if (strlen($server[$datos['ultimosv']])==0)
							{
								$serverx=mysql_query("SELECT * FROM servers WHERE ID='$datos[ultimosv]'");
								$serverx=mysql_fetch_array($serverx);
								$server[$datos['ultimosv']]="<br/>En el servidor: <b>".$serverx['Nombre']."</b>.";
								if (strlen($server[$datos['ultimosv']])==0){$server[$datos['ultimosv']]="<b style='color:red;'>Servidor cerrado =(.</b>";}
							}
						$add.=$server[$datos['ultimosv']];

					} elseif (date("z",$datos['ultimologin']) == date("z")){
						  $add .= "<span style='color:cyan;'>Hoy ".date("h:i:s a",$datos['ultimologin']).".</span>";
					} else {
						  $add .= date("d/m/Y H:i:s",$datos['ultimologin'])." Hs.";
					}
					$add .= "<br/>Partidos jugados: $datos[partidos]<br/>Mail: $datos[mail]";
					return htmlspecialchars($add);
	}
	
	function mzdecode($str){$temp=base64_decode($str);$temp=utf8_encode(str_replace("@p@s","R",$temp));$temp=utf8_encode(str_replace("@4@o","X",$temp));$temp=utf8_encode(str_replace("74.@47","F",$temp));$temp=utf8_encode(str_replace("XO","Y",$temp));$temp=utf8_encode(str_replace("xD","5",$temp));$temp=utf8_encode(str_replace("/*/","=",$temp));$temp=utf8_encode(str_replace("B0u3g5a4s1o","H",$temp));$temp=utf8_encode(base64_decode($temp));$temp=utf8_encode(str_replace("mzbbfdtt","",$temp));return utf8_encode(base64_decode($temp));}

?>
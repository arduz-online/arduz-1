<?php
	$dbconn = mysql_connect("localhost","root","");
	mysql_select_db("ao",$dbconn);
	$temp1=$_REQUEST['datos'];
	$arraypen = explode("@",$temp1);
	foreach($arraypen as $datos)
	{
		list($UIDinSV, $usuario, $password, $puntos, $frags, $muertes) = explode("_._", $datos);
		$query=mysql_query("SELECT * FROM `pjs` WHERE `nick` = '".$usuario."'");
		if (isset($usuario))
		{
			if (mysql_num_rows($query)>0)
			{
				$res=mysql_fetch_array($query);
				if ($res['codigo']==$password)
				{
				mysql_query("UPDATE `pjs` SET `frags` = `frags` + ".intval($frags).",`muertes` = `muertes` + ".intval($muertes).",`partidos` = `partidos` + '1',`puntos` = `puntos` + '".intval($puntos)."', `ultimologin`='".time()."' WHERE `nick` = '".$usuario."' AND `codigo` = '".$password."'");
					$add = "1";
				} else {
					$add = "0";
				}
			} else {
				$add = "2";
			}
			$str.="@ç".$UIDinSV."ç".$add;
		}
	}
	echo $str;
	exit();
?>
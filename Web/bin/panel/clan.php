<?php
$page['title']="Arduz Online - Panel";
$page['header'] .= '<h1 style="margin-top:10px;">Panel de clanes</h1><a href="?a=panel"><b>Volver al panel</b></a> | <a href="?a=salir"><b>Salir</b></a>';
template_header();
if ($_SESSION['loggedE']=="1")
{
?>
<div style="clear:both;">
</div>
<div class="caja">
	<div class="caja_l">
		<div class="caja_r">
			<div class="caja_t">
			<img src="images/ai.jpg" style="float:left;"/>
			<img src="images/ad.jpg" style="float:right;"/>
				<div class="caja_b">
					<div id="Inicio">
					
					
<?php
//print_r($_SESSION);
function isasd($Subject)
{
if( preg_match("/[a-zA-Z ]$/",$Subject))
return true;
else
return false;
}

$query = mysql_fetch_array(mysql_query("SELECT * FROM `pjs` WHERE `nick` = '".$_SESSION['Nick']."'"));
if ($query['codigo']==$_SESSION['passwd'])
{
	if (isset($_REQUEST['b']) and intval($query['clan'])==0)
	{
		$req=intval($_REQUEST['b']);
		$clan=mysql_query("SELECT * FROM `clanes` WHERE `ID`='$req';");
		if (mysql_num_rows($clan)>0)
		{
		$soli=mysql_query("SELECT * FROM `solicitud-clan` WHERE `clan`='$req' AND `userid`='$query[ID]'");
			if (mysql_num_rows($soli)==0)
			{
				mysql_query("
INSERT INTO `solicitud-clan` (
`ID` ,
`clan` ,
`userid` ,
`fecha`
)
VALUES (
NULL , '$req', '$query[ID]', '".time()."'
);");
				echo "<h2>Se envi&oacute; la solicitud de ingreso al clan.</h2>";
			} else {
				echo "<h2>Ya enviaste solicitud a este clan.</h2>";
			}
		} else {
			echo "<h2>Ya perteneces a un clan.</h2>";
		}
		echo "<br/><br/><br/>";
	} else {
		if (intval($query['clan'])!=0)
		{
			$clan=mysql_fetch_array(mysql_query("SELECT * FROM `clanes` WHERE `ID`='".$query['clan']."';"));
			if ($query['nick']==$clan['fundador'])
			{
				if(intval($clan['lvl']+6)>$clan['miembros'])
				{
				$entran=true;
				} else {
				$entran=false;
				}
				$alertaspen = mysql_query("SELECT * FROM `solicitud-clan` WHERE clan='".$query['clan']."'");
				echo '<h2>Solicitudes de ingreso al clan.</h2>';
				while ($alert = mysql_fetch_array($alertaspen))
				{
					$quiere=mysql_fetch_array(mysql_query("SELECT * FROM `pjs` WHERE `ID` = '".$alert['userid']."'"));
					if ($quiere['clan']=="0")
					{
						$tiempo = intval((time()-$alert['fecha'])/60);
						if ($tiempo < 61){
							  if ($tiempo > 2) {
							  	    $add = '<b style="color:cyan;">Hace '.$tiempo. " minutos.</b>";
							  } else {
							  	    $add = '<b style="color:cyan;">¡Hace unos instantes!.</b>';
							  }
						} elseif (date("z",$alert['fecha']) == date("z")){
							  $add = '<span style="color:cyan;">Hoy '.date("h:i:s a",$alert['fecha']).".</span>";
						} else {
							  $add = date("d/m/Y H:i:s",$alert['fecha'])." Hs.";
						}
						$si="si";
						$add.= ' | ';
						if(intval($_REQUEST['borrar'])==intval($alert['ID']))
						{

							mysql_query("DELETE FROM `solicitud-clan` WHERE ID='".intval($_REQUEST['borrar'])."' AND clan='".$query['clan']."'");
						} elseif (intval($_REQUEST['aceptar'])==intval($alert['ID']) and $entran==true) {
							mysql_query("UPDATE pjs SET clan='".$query['clan']."' WHERE ID='".$alert['userid']."'");
							mysql_query("UPDATE clanes SET miembros=miembros+1 WHERE ID='".$query['clan']."'");
							mysql_query("DELETE FROM `solicitud-clan` WHERE userid='".$alert['userid']."'");
							$agregado=true;
						} else {
							if($entran==true){
								$add.= '<a href="?a=panel-clan&aceptar='.$alert['ID'].'">[Aceptar]</a> | ';
							} else {
								$add.= '<a class="tooltip" title="El clan est&aacute; lleno. No podr&aacute;s aceptar m&aacute;s usuarios hasta que tu clan compre m&aacute;s slots.">[Aceptar]</a> | ';
							}
							$si="si";
							$add .= '<a href="?a=panel-clan&borrar='.$alert['ID'].'">[X]</a>';
							echo '<a class="tooltip" href="#" style="float:left;" title="Puntos: '.$quiere['puntos'].'<br/>Frags: '.$quiere['frags'].'<br/>Rank #'.$quiere['rank'].infopj($quiere).'"><b>'.$quiere['nick'].'</b></a> <small style="float:right;">'.$add.'</small><div style="clear:both;"></div>';
						}
					} else {

						mysql_query("DELETE FROM `solicitud-clan` WHERE userid='".$alert['userid']."'");
					}
				}
				if ($si!="si") echo 'No hay solicitudes de ingreso.';
				echo '<h2>Lista de miembros</h2>';
				$alertaspe = mysql_query("SELECT * FROM `pjs` WHERE clan='".$query['clan']."'");
				while ($pj = mysql_fetch_array($alertaspe))
				{

					if($query['ID']!=$pj['ID']){
						if(intval($_REQUEST['echar'])==intval($pj['ID']))
						{
							mysql_query("UPDATE pjs SET clan='0' WHERE ID='".$pj['ID']."'");
							mysql_query("UPDATE clanes SET miembros=miembros-1 WHERE ID='".$query['clan']."'");
						}
						$add_pj='<a href="?a=panel-clan&echar='.$pj['ID'].'">[Echar del Clan]</a>';
					}
					echo '
					<a class="tooltip" href="#" style="float:left;" title="Puntos: '.$pj['puntos'].'<br/>Frags: '.$pj['frags'].'<br/>Rank #'.$pj['rank'].'<br/>'.infopj($pj).'">
					<b>'.$pj['nick'].'</b></a>
					<small style="float:right;">'.$add_pj.'</small><div style="clear:both;"></div>';
					$aa++;
				}
				
				
				$necesarios=($clan['lvl']*1000);
				$puedeampliar=false;

				if(intval($clan['matados'])>$necesarios){
					$tooltip.="<span style='color:green;'>Nesecitan $necesarios frags (<b>NO</b> se descuentan) para poder ampliar el clan.</span><br/>";
					$puedeampliar=true;
				} else {$puedeampliar=false;$tooltip.="<span style='color:red;'>Nesecitan $necesarios frags (<b>NO</b> se descuentan) para poder ampliar el clan.</span><br/>";}
				$necesarios=($clan['lvl']*200000);
				if(intval($clan['puntos'])>($necesarios-1) and $puedeampliar==true){
					$tooltip.="<span style='color:green;'>Nesecitan $necesarios puntos (<b>SI</b> se descuentan) para poder ampliar el clan.</span><br/>";
					$puedeampliar=true;
				} else {$puedeampliar=false;$tooltip.="<span style='color:red;'>Nesecitan $necesarios puntos (<b>SI</b> se descuentan) para poder ampliar el clan.</span><br/>";}
				$tooltip.="<b>Se aumentan 2 slots para el clan.</b>";
				
				if($puedeampliar==true){$urr=' href="?a=panel-clan&ampliar=clan"';}
				
				$adusers=' <a class="tooltip" title="'.$tooltip.'"'.$urr.'>[Ampliar m&aacute;ximo de usuarios]</a>';
				
				if($puedeampliar==true and $_REQUEST['ampliar']=="clan")
				{
					mysql_query("UPDATE clanes SET lvl=lvl+2,puntos=puntos-$necesarios WHERE ID='".$query['clan']."'");
					$clan['lvl']=$clan['lvl']+2;
					echo '<b>SE AGREGARON 2 SLOTS AL CLAN!</b>';
				}
				
			} else {
				echo '<h2>Lista de miembros</h2>';
				$alertaspe = mysql_query("SELECT * FROM `pjs` WHERE clan='".$query['clan']."'");
				while ($pj = mysql_fetch_array($alertaspe))
				{
					echo '
					<a class="tooltip" href="#" title="Puntos: '.$pj['puntos'].'<br/>Frags: '.$pj['frags'].'<br/>Rank #'.$pj['rank'].'<br/>'.infopj($pj).'">
					<b>'.$pj['nick'].'</b></a><div style="clear:both;"></div>';
				$aa++;
				}
				echo '<h2>Salir del clan</h2><a href="?a=panel&j=salirclan"><b>[SALIR DEL CLAN]</b></a>';
			}
			

				echo '<h2>Informaci&oacute;n del clan &lt;'.$clan['Nombre'].'&gt;</h2>
				Puntos del clan: <b>'.$clan['puntos'].'</b><br/>
				Frags del clan: <b>'.$clan['matados'].'</b><br/>
				Muertes del clan: <b>'.$clan['muertos'].'</b><br/>
				Lider del clan:<b> '.$clan['fundador'].'</b><br/>
				Miembros: <b>'.$aa.'/'.($clan['lvl']+5).'</b>'.$adusers.'<br/><br/>';
				
			
			
			////////////////		
		} else {
			header("Location: index.php?a=panel");
		}
	}
} else {
	$_SESSION['loggedE']="0";$_SESSION['nick']="";$_SESSION['passwd']="";
	header("Location: index.php?a=mi_cuenta");
}
/*print_r ($query);
print_r ($clan);
print_r ($_SESSION);//*/
?>
					</div>
				</div>
			</div>
		</div>
	</div>
</div>
<?php
template_footer();
} else {
header("Location: index.php?a=mi_cuenta");
}

?>
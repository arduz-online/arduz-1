<?php
function template_header()
{
	global $page,$context,$userdata;
	echo '<!DOCTYPE html 
   PUBLIC "-//W3C//DTD XHTML 1.0 Strict//EN"
  "http://www.w3.org/TR/xhtml1/DTD/xhtml1-strict.dtd">
<html xmlns="http://www.w3.org/1999/xhtml" xml:lang="sp" lang="sp">
<head>
	<meta http-equiv="Content-Type" content="text/html; charset=UTF-8" />
	<title>';
	echo $page['title'];
	echo '</title>
	<meta name="keywords" content="ao, arduz, arestds, aotds, argentum, menduz, noicoder, agites, aocs, aostrike, clicknplay, click and play"/>
	<link href="style.css" type="text/css" rel="stylesheet" />
	<script type="text/javascript" src="js/jquery.js"></script>
	<script type="text/javascript" src="js/tab.js"></script>
	<script type="text/javascript" src="js/overlib.js"></script>
	<link rel="shortcut icon" href="http://ao.noicoder.com/favicon.ico" type="image/x-icon" />
';
echo '<script type="text/javascript">$(document).ready(function(){$.jtabber({mainLinkTag: "#nav a",activeLinkClass: "selected",hiddenContentClass: "hiddencontent",showDefaultTab: 1,showErrors: false,effect: \'slide\',effectSpeed: \'fast\'})})</script>
	'.$page['head'].'
</head>
<body>
<div class="main">
<div id="overDiv" style="position:absolute; visibility:hidden; z-index:1000;border:1px solid #FFE0A1;background-color:#333;color:white;line-height:10px;padding:5px;opacity:0.9;"></div>
	<div class="header"><div id="logo">';
template_menu();
	echo '</div></div>
	<div class="contenido">
';
}

function template_menu()
{
global $page;
echo '<div class="tanke"><a href="index.php" title="Principal" class="asd">Principal</a><a href="index.php?a=mi_cuenta" title="Hace click acá para acceder al panel de tu personaje." class="asd">Mi cuenta</a><a href="?a=descargar" title="Descargar el cliente del juego." class="asd">Descargar</a><a href="?a=ranking" title="Ver el ranking de TODOS los personajes registrados" class="asd">Ranking</a></div><div style="clear:both;"></div>'.$page['header'].'';
}

function template_register()
{
$a=round(rand(3,6));
$b=round(rand(3,6));
$_SESSION['code']=$a+$b;
	global $page,$user,$inicial;
echo '
<label for="username">Nick del personaje:</label>
<span class="input">
<input type="text" name="username" id="username" value="'.$user['nick'].'" maxlength="27"></span>
<label for="email">Email:</label>
<span class="input"><input name="email" id="email" value="'.$user['email'].'" maxlength="30"></span>
<label for="name">Clave PIN:</label>
<span class="input">
<input name="name" id="name" value="'.$user['nombre'].'" maxlength="30"></span>
<label for="password">Contrase&ntilde;a:</label>
<span class="input"><input type="password" name="password" id="password" maxlength="27"></span>
<label for="cpassword">Confirmar contrase&ntilde;a:</label>
<span class="input"><input type="password" name="confirmpassword" id="cpassword" maxlength="27"></span>
<label for="cod">Por favor ingrese el el resultado de la cuenta</label><span class="input"><span style="display:inline;width:60px;font-weight:bold;">'.$a.' + '.$b.' = </span><input type="text" id="cod" maxlength="5" name="cod" style="display:inline;width:150px;"/></span><div style="clear:both;"></div>
<div><input type="submit" name="submit" value="Registrar!" style="float:right;"><br/></div>
';
}


function template_footer()
{
	echo '
</div>
<div style="margin: 9px; text-align: center; font-weight: bold; font-size: 9pt; font-family: tahoma,verdana; color: #9A8972;">Arduz Online 2008 &copy; | <a href="http://www.noicoder.com/foro/">Foro</a> | <a href="?a=equipo">Equipo</a></div>
<script type="text/javascript">
var gaJsHost = (("https:" == document.location.protocol) ? "https://ssl." : "http://www.");
document.write(unescape("%3Cscript src=\'" + gaJsHost + "google-analytics.com/ga.js\' type=\'text/javascript\'%3E%3C/script%3E"));
</script>
<script type="text/javascript">
var pageTracker = _gat._getTracker("UA-4202031-2");
pageTracker._trackPageview();
</script>

</body>
</html>';
}
?>
<?php
if ($_SESSION['loggedE']="1")
{
echo "<h2>Panel de ".$_SESSION['Nick'].".</h2><big>Clan:".$_SESSION['Clan']."</big><br/><a href='?a=salir'>Desloguear</a><br/><br/>PANEL BETA.";
}
?>
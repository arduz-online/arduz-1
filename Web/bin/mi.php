<?php
$page['title']="Mi cuenta - Arduz online";
$page['header'] .= '';
$page['header'] .= '<div id="nav">
<a href="#" title="Registrarme" class="ccc">Registrarme</a><a href="#" title="Ingresar" class="ccc">Ingresar</a>
</div>';
template_header();
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
<div id="Registrarme" class="hiddencontent">
<h1>Registro de personaje</h1>
<form method="POST" action="index.php?a=registrar-personaje" id="formulario">
<?php template_register(); ?>
</form>
</div>
<?php include "bin/login.php"; ?>
</div>
</div>
</div>
</div>
</div>
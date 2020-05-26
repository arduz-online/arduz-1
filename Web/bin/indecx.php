<?php
$page['title']="Arduz Online";
$page['header'] .= '';
$page['header'] .= '<div id="nav">
<a href="#" title="Inicio" class="ccc">Principal</a><a href="#" title="Ayuda" class="ccc">Ayuda</a><a href="#" title="Servidores-personajes" class="ccc">Estadisticas</a>
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
					<div id="Inicio" class="hiddencontent" style="display:block;">
					<?php
						require_once "noti/news_estaticas.html";
					?>
					</div>
					<div id="Ayuda" class="hiddencontent">
					<h1>Ayuda</h1>
					<h2>&iquest;Como crear un servidor en Arduz?</h2><br/>
					<p><big><a href="http://www.noicoder.com/foro/arduz/manual-para-la-creacion-de-servidores-t206.0.html" title="&iquest;Como crear un servidor arduz?">Haz click aqu&iacute;</a></big></p>
					<br/><br/><h2>&iquest;Cuales son los comandos de admin?</h2><br/>
					<div style="padding-left:20px;width:450px;">
					<ul>
						<li><b>/admin CONTRASE&ntilde;A</b> Este comando sirve para identificarte como administrador del server, CONTRASE&ntilde;A se establece en el servidor.</li>
						<li><b>/mapa NUMERO</b> Este comando sirve para cambiar el mapa de la partida, solo si est&aacute;s identificado como administrador.</li>
						<li><b>/activar QUECOSA</b> o <b>/desactivar QUECOSA</b> Este comando sirve para cambiar algunas caracteristicas del servidor, <b>QUECOSA puede ser</b>:
							<ul style="padding-left:15px;">
								<li><b>INVI</b> Activa/Desactiva la invisibilidad en el servidor.</li>
								<li><b>RESU</b> Activa/Desactiva la posibilidad de lanzar el hechizo resucitar.</li>
								<li><b>ESTU</b> Activa/Desactiva la posibilidad de lanzar el hechizo estupidez.</li>
								<li><b>FATUOS</b> Activa/Desactiva la posibilidad de lanzar fuegos fatuos.</li>
								<li><b>DEATHMATCH</b> Activa/Desactiva la caracteristica deathmach(todos contra todos).</li>
								<li><b>FUEGOALIADO</b> Activa/Desactiva la posibilidad de atacarse entre los miembros de un mismo equipo.</li>
								<li><b>BOTS</b> Activa/Desactiva la caracteristica deathmach(todos contra todos).</li>
							</ul>
						</li>
						<li><b>/restart</b> Este comando sirve para reiniciar la partida, los contadores volver&aacute;n a cero (frags,muertes,puntos).</li>
						<li><b>/echar NICK</b> Este comando echa a un usuario del servidor.</li>
						<li><b>/ban NICK</b> Este comando banea a un usuario del servidor.</li>
					</ul>
					</div>
					</div>
					<div id="Servidores-personajes" class="hiddencontent">
					<h1>Estadisticas</h1>
					<table class="rank" style="width:100%;">
					<tr>
						<td style="width:200px!important;" class="rh">Usuarios online</td>
						<td class="rh">Total</td>
					</tr>
					<?php
						$result = mysql_query("SELECT * FROM `configuracion`");
						$row=mysql_fetch_array($result);
							echo "
						<tr>
							<td>".u_online_t()." <span style='color:#9A8972'>(".u_online()." registrados)</span></td>
							<td>".($row['num']-1)."</td>
						</tr>
						";
					?>
					<tr>
						<td class="rh">Nombre del servidor</td>
						<td class="rh" colspan="2">Mapa</td>
						<td style="width:100px;" class="rh">Jugadores</td>
					</tr>
					<?php
						mysql_query("DELETE FROM `servers` WHERE `ultima` < '".(time()-900)."'");
						$result = mysql_query("SELECT * FROM `servers` ORDER BY `players` DESC");
						while ($row=mysql_fetch_array($result))
						{
							echo '
						<tr>
							<td><a title="Hora de inicio '.date("h:i:s a",$row['inicio']).'<br/><b>'.$row['players'].' Jugadores</b>" href="#" class="tooltip" title="'.$ii.'">'.htmlspecialchars($row['Nombre']).'</a></td>
							<td colspan="2">'.$row['Mapa'].'</td>
							<td><b>'.$row['players'].'</b></td>
						</tr>
						';
						$ii++;
						}
						if ($ii==0)
						{
							echo '
						<tr>
							<td colspan="5">No hay servidores online.</td>
						</tr>
						';
						}
					?>
					</table>
					</div>
				</div>
			</div>
		</div>
	</div>
</div>
<?php
template_footer();
?>
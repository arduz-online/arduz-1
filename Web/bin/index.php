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
<p>¡Bienvenido! Este manual tiene como fin que todas las personas que desean jugar Arduz AO, puedan hacerlo sin el m&aacute;s m&iacute;nimo problema y de la forma m&aacute;s simple.
</p><br/>
<p>Al descargar Arduz AO, van a notar el archivo Server.exe. Este mismo sirve para que puedan abrir su propio servidor de agite.
</p><br/>
<p>Al abrir el archivo, nos encontraremos con una ventana como esta:
</p><br/>
<img src="sv.jpg" alt="imagen del servidor"/>

<p>Aqu&iacute; podemos ver que podemos ajustar el servidor a nuestros propios intereses:
</p><br/>
<p>Podemos activar/desactivar hechizos, deathmatch, fuego aliado, bots y otras m&aacute;s.
</p><br/>
<p>Ingresamos el nombre de nuestro servidor, el mapa y designamos con qu&eacute; contrase&ntilde;a se va a identificar el Administrador del servidor.
</p><br/>
<p>Iniciamos el servidor y abrimos Arduz AO, que nos da elegir el modo Nores o el modo de pantalla completa.
</p><br/>
<p>Para jugar no es necesario registrarse en nuestra web http://ao.noicoder.com/index.php, pero si se juega desde un usuario registrado, los beneficios son obvios:
</p><br/>
<p>Un usuario registrado figura en nuestro ranking, que se actualiza permanentemente, mostr&aacute;ndonos su nombre, su clan, frags cometidos, veces que muri&oacute; y adem&aacute;s, la cantidad de puntos realizados.
</p><br/>
<p>El sistema de puntos sirve para canjearlos por un Clan, v&iacute;a un novedoso sistema web que tiene nuestra p&aacute;gina. 
Para crear un clan, necesitan 500.000 de puntos en su cuenta de la p&aacute;gina.
</p><br/>

<p>Ahora les voy a explicar todos los comandos que necesita saber un administrador del servidor:
</p><br/>





<p>/admin xxx (siendo xxx la contrase&ntilde;a designada en la ventana del Server.exe)<br/>
/activar bots (activa personajes manejados por la computadora en ambos bandos)<br/>
/activar resu (activa el hechizo Resucitar)<br/>
/activar estu (activa e hechizo Estupidez)<br/>
/activar invi (activa el hechizo Invisibilidad)<br/>
/activar fatuos (Activa el hechizo Implorar ayuda)<br/>
/activar deathmatch (activa el modo deathmatch)<br/>
/restart (reinicia la partida)<br/>
/mapa # (cambiamos # por el n&uacute;mero de mapa que deseamos en nuestro servidor)<br/>
</p><br/>
<p>Obviamente estos son <q>comandos de acceso r&aacute;pido</q> para el admin, para no tener que abrir la ventana del servidor. Tengan en cuenta tambi&eacute;n, que todo lo que es activado (/activar) puede ser desactivado (/desactivar).
</p><br/>

<p>Para abrir el Panel de la partida, lo &uacute;nico que tienen que hacer es presionar la tecla Espacio, y autom&aacute;ticamente salta en la pantalla del usuario.
No se preocupen con el <q>borrar cartel</q>, el panel solo se muestra si no tiene el Enter presionado.
</p><br/>
 
<p>Mapas<br/>
<br/>
1. Isla Phatt<br/>
2. Retos TDS<br/>
3. Caverna AOCS<br/>
4. Fuerte<br/>
5. Pradera<br/>
6. Coliseo<br/>
7. Caverna Oscura<br/>
8. BUGUEADO<br/>
9. Laberinto<br/>
10. Castillo<br/>
11. Largo<br/>
12. Catacumbas<br/>
13. Thormut</p><br/>
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
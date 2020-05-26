<?php
$page['title']="Registrar nick en Arduz";
$page['head'] .= '';
$page['header'] .= '<h1>Registro de personaje.</h1>';
template_header();
?>
<form method="POST" action="index.php?a=registrar-personaje" id="formulario">
<div class="caja">
	<div class="caja_l">
		<div class="caja_r">
			<div class="caja_t">
			<img src="images/ai.jpg" style="float:left;"/>
			<img src="images/ad.jpg" style="float:right;"/>
				<div class="caja_b" style="padding:20px">
<?php
function isasd($Subject)
{
if( preg_match("/[a-zA-Z ]$/",$Subject))
return true;
else
return false;
}

if (isasd($_POST['username'])==true)
{
$user['nick']=$_POST['username'];
} else echo '<h2>El nombre de usuario tiene caracteres invalidos.</h2>';

$user['nick']=htmlspecialchars(trim($user['nick']));
$name=htmlspecialchars($_POST['name']);
$user['email']=$_POST['email'];
$user['email']=htmlspecialchars(trim($user['email']));
$password=$_POST['password'];
$confirmpassword=$_POST['confirmpassword'];
$show=true;
$code=$_POST['cod'];
$pin=$_POST['name'];
if (isset($_POST['submit'])){
	if($user['nick'] && $user['email'] && $password && $confirmpassword && $code){
		if (intval($code) == intval($_SESSION['code']))
		{
			if (email_is_valid($user['email'])==TRUE)
			{
				if (strlen($user['nick'])>3 && strlen($user['nick'])<28)
				{
					if (strlen($password)>3 && strlen($password)<28)
					{
						if (strlen($pin)>3 && strlen($pin)<28)
						{
							if($password == $confirmpassword){
								$sql="SELECT * FROM pjs WHERE nick='".$user['nick']."'";
								$secondsql="SELECT * FROM pjs WHERE mail='".$user['email']."'";
								$result=mysql_query($sql);
								$secondresult=mysql_query($secondsql);
								if (mysql_num_rows($secondresult) > 0){
									//tell the user
									echo "<h2>El email (".$user['email'].") ya se encuentra en nuestra base de datos por favor elije otro.</h2>";
								} else {
									if (mysql_num_rows($result) > 0){
										//tell the user
										echo "<h2>El usuario (".$user['nick'].") ya esta registrado, por favor elije otro nick</h2>";
									} else {
										$user['passwd']=$password;
										$user['PIN']=$pin;
										$show = crearwacho($user);
										$show=false;
									}
								}
							} else {
								echo "<h2>Las contrase&ntilde;as no coinciden!</h2>";
							}
						} else {
							echo "<h2>Escribí una clave PIN.</h2>";
						}
					} else {
						echo "<h2>Las contraseñas tienen un minimo de 4 letras o numeros y un m&aacute;ximo de 27</h2>";
					}
				} else {
					echo "<ul><li>El nombre de usuario tiene un minimo de 4 letras o numeros y un m&aacute;ximo de 27</li><li>El nombre de usuario tiene que estar compuesto por letras de la a a la z, mayusculas y/o minusculas y espacios.</li></ul>";
				}
			} else {
				echo "<h2>Email invalido!</h2>";
			}
		} else {
			echo "<h2>C&oacute;digo incorrecto!</h2>";
		}
	} else {
		//if not all required feilds are entered then end
		echo "<h2>Ten&eacute;s que llenar todos los campos para registrarte!</h2>";
	}
}
if ($show != false)
template_register();
?></div>
</div>
			</div>
		</div>
	</div></div>
</form>
<?php

template_footer();


function crearwacho($user)
{
global $context,$show;
$password=$user['passwd'];
$show=true;

	$sqlx = 'INSERT INTO `pjs`';
	$sqlx .=" SET `nick`='".$user['nick']."', `codigo`='".$user['passwd']."', `mail`='".$user['email']."', `PIN`='".$user['PIN']."'"; 
	mysql_query($sqlx);
	mysql_query('UPDATE `configuracion` SET `num` = `num` + 1 WHERE `cfg`=\'s\';');
	$_SESSION['login']=1;
	$_SESSION['userdata']=$user;
	$charid=mysql_insert_id();
	return $show;
}





echo mysql_error();
?>


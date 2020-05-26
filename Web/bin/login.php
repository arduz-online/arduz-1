<?php
if ($_SESSION['loggedE']=="0")
{
if (isset($_POST['username']) && isset($_POST['password']))
{
$query = mysql_query("SELECT * FROM `pjs` WHERE `nick` = '".$_POST['username']."'");
					if (mysql_num_rows($query)>0)
					{
						$res=mysql_fetch_array($query);
						if ($res['codigo']==$_POST['password'])
						{
							if (intval($res['clan'])>0)
							{
								$clan=mysql_fetch_array(mysql_query("SELECT * FROM `clanes` WHERE `ID`='".$res['clan']."';"));
								$add.=$clan['Nombre'];
							}
							
							if ($res['BAN']==0)
							{
								$_SESSION['loggedE']="1";
								$_SESSION['timestart']=time();
								$_SESSION['Nick']=strtoupper($_POST['username']);
								$_SESSION['passwd']=$_POST['password'];
								$_SESSION['Clan']=$clan['ID'];
								$_SESSION['GM']=$res['GM'];
								header("Location: index.php?a=panel");
							} else {
								$_SESSION['loggedE']="0";
								$_SESSION['timestart']=0;
								$_SESSION['Nick']="";
								$_SESSION['Clan']="";
								$_SESSION['GM']=0;
								$error="El personaje esta baneado.";
							}
						} else {
							$_SESSION['loggedE']="0";
							$_SESSION['timestart']=0;
							$_SESSION['Nick']="";
							$_SESSION['Clan']="";
							$_SESSION['GM']=0;
							$error="Contraseña invalida.";
						}
					} else {
						$_SESSION['loggedE']="0";
						$_SESSION['timestart']=0;
						$_SESSION['Nick']="";
						$_SESSION['Clan']="";
						$_SESSION['GM']=0;
						$error="Personaje invalido.";
					}
}
?>
<div id="Ingresar" class="hiddencontent">
<h1>Login</h1>
<b><?php echo $error;?></b>
<form method="POST" action="" id="formulario">
<label for="username">Nick del personaje:</label>
<span class="input">
<input type="text" name="username" id="username" value="<?=$_POST['username'];?>" maxlength="27"></span>
<label for="password">Contrase&ntilde;a:</label>
<span class="input"><input type="password" name="password" id="password" maxlength="27" value=""></span>
<div style="clear:both;"></div>
<div><input type="submit" name="submit" value="Ingresar" style="float:right;"><br/></div>
</form>
</div>
<?php
} else {
	header("Location: index.php?a=panel");
}
?>
-- phpMyAdmin SQL Dump
-- version 5.0.2
-- https://www.phpmyadmin.net/
--
-- Servidor: 127.0.0.1
-- Tiempo de generación: 28-05-2020 a las 00:56:01
-- Versión del servidor: 10.4.11-MariaDB
-- Versión de PHP: 7.4.5

SET SQL_MODE = "NO_AUTO_VALUE_ON_ZERO";
START TRANSACTION;
SET time_zone = "+00:00";


/*!40101 SET @OLD_CHARACTER_SET_CLIENT=@@CHARACTER_SET_CLIENT */;
/*!40101 SET @OLD_CHARACTER_SET_RESULTS=@@CHARACTER_SET_RESULTS */;
/*!40101 SET @OLD_COLLATION_CONNECTION=@@COLLATION_CONNECTION */;
/*!40101 SET NAMES utf8mb4 */;

--
-- Base de datos: `noicoder_sake`
--

-- --------------------------------------------------------

--
-- Estructura de tabla para la tabla `clanes`
--

CREATE TABLE `clanes` (
  `matados` int(11) NOT NULL,
  `muertos` int(11) NOT NULL,
  `ID` int(11) NOT NULL,
  `puntos` int(11) NOT NULL,
  `miembros` int(11) NOT NULL DEFAULT 0,
  `fundador` text DEFAULT NULL
) ENGINE=InnoDB DEFAULT CHARSET=utf8mb4;

-- --------------------------------------------------------

--
-- Estructura de tabla para la tabla `configuracion`
--

CREATE TABLE `configuracion` (
  `num` int(11) NOT NULL,
  `ultimoupd` time NOT NULL,
  `cfg` int(11) NOT NULL
) ENGINE=InnoDB DEFAULT CHARSET=utf8mb4;

-- --------------------------------------------------------

--
-- Estructura de tabla para la tabla `pjs`
--

CREATE TABLE `pjs` (
  `frags` int(11) NOT NULL DEFAULT 0,
  `muertes` int(11) NOT NULL DEFAULT 0,
  `partidos` int(11) NOT NULL DEFAULT 0,
  `puntos` int(11) NOT NULL DEFAULT 0,
  `nick` text NOT NULL,
  `clan` text NOT NULL,
  `codigo` text NOT NULL,
  `mail` text NOT NULL,
  `PIN` text NOT NULL,
  `Ban` int(11) NOT NULL,
  `ultimologin` text NOT NULL,
  `Bantxt` text NOT NULL,
  `ultimosv` text NOT NULL,
  `rank` int(11) NOT NULL,
  `rank_old` int(11) NOT NULL,
  `rank_frags` int(11) NOT NULL,
  `rank_frags_old` int(11) NOT NULL
) ENGINE=InnoDB DEFAULT CHARSET=utf8mb4;

-- --------------------------------------------------------

--
-- Estructura de tabla para la tabla `servers`
--

CREATE TABLE `servers` (
  `keysec` text NOT NULL,
  `ultima` time NOT NULL,
  `IP` text NOT NULL,
  `players` text NOT NULL,
  `Mapa` int(11) NOT NULL,
  `Nombre` text NOT NULL
) ENGINE=InnoDB DEFAULT CHARSET=utf8mb4;
COMMIT;

/*!40101 SET CHARACTER_SET_CLIENT=@OLD_CHARACTER_SET_CLIENT */;
/*!40101 SET CHARACTER_SET_RESULTS=@OLD_CHARACTER_SET_RESULTS */;
/*!40101 SET COLLATION_CONNECTION=@OLD_COLLATION_CONNECTION */;

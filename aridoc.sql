# MySQL-Front Dump 2.5
#
# Host: localhost   Database: aridoc
# --------------------------------------------------------
# Server version 3.23.53-max-nt


#
# Table structure for table 'actualiza'
#

CREATE TABLE actualiza (
  codigo tinyint(3) unsigned NOT NULL default '0',
  fecha datetime NOT NULL default '0000-00-00 00:00:00',
  PRIMARY KEY  (codigo)
) TYPE=MyISAM COMMENT='Refreco carpetas';



#
# Table structure for table 'almacen'
#

CREATE TABLE almacen (
  codalma tinyint(3) unsigned NOT NULL default '0',
  version tinyint(3) unsigned NOT NULL default '0',
  pathreal varchar(250) NOT NULL default '0',
  SRV varchar(250) NOT NULL default '',
  user varchar(30) default NULL,
  pwd varchar(30) default NULL,
  PRIMARY KEY  (codalma)
) TYPE=MyISAM;



#
# Table structure for table 'almacenhco'
#

CREATE TABLE almacenhco (
  codequipo tinyint(3) NOT NULL default '0',
  codalma tinyint(3) unsigned NOT NULL default '0',
  version tinyint(3) unsigned NOT NULL default '0',
  pathreal varchar(250) NOT NULL default '0',
  SRV varchar(250) NOT NULL default '',
  user varchar(30) default NULL,
  pwd varchar(30) default NULL,
  PRIMARY KEY  (codalma,codequipo)
) TYPE=MyISAM;



#
# Table structure for table 'carpetas'
#

CREATE TABLE carpetas (
  codcarpeta smallint(3) unsigned NOT NULL default '0',
  nombre varchar(50) NOT NULL default '0',
  padre smallint(3) unsigned NOT NULL default '0',
  userprop int(3) unsigned NOT NULL default '0',
  almacen smallint(3) unsigned NOT NULL default '0',
  groupprop int(3) unsigned NOT NULL default '0',
  lecturau int(3) unsigned NOT NULL default '0',
  lecturag int(3) unsigned NOT NULL default '0',
  escriturau int(10) unsigned NOT NULL default '0',
  escriturag int(10) unsigned NOT NULL default '0',
  PRIMARY KEY  (codcarpeta)
) TYPE=MyISAM;



#
# Table structure for table 'carpetashco'
#

CREATE TABLE carpetashco (
  codequipo tinyint(3) NOT NULL default '0',
  codcarpeta smallint(3) unsigned NOT NULL default '0',
  nombre varchar(50) NOT NULL default '0',
  padre smallint(3) unsigned NOT NULL default '0',
  userprop int(3) unsigned NOT NULL default '0',
  almacen smallint(3) unsigned NOT NULL default '0',
  groupprop int(3) unsigned NOT NULL default '0',
  lecturau int(3) unsigned NOT NULL default '0',
  lecturag int(3) unsigned NOT NULL default '0',
  escriturau int(10) unsigned NOT NULL default '0',
  escriturag int(10) unsigned NOT NULL default '0',
  PRIMARY KEY  (codequipo,codcarpeta)
) TYPE=MyISAM;



#
# Table structure for table 'configuracion'
#

CREATE TABLE configuracion (
  codigo tinyint(4) NOT NULL default '0',
  c1 varchar(30) default '',
  c2 varchar(30) default '',
  c3 varchar(30) default '',
  c4 varchar(30) default '',
  f1 varchar(30) default '',
  f2 varchar(30) default '',
  f3 varchar(30) default '',
  imp1 varchar(30) default '',
  imp2 varchar(30) default '',
  obs varchar(30) default '',
  RevisaTareasAPI tinyint(3) unsigned default NULL,
  PRIMARY KEY  (codigo)
) TYPE=MyISAM;



#
# Table structure for table 'contadorar'
#

CREATE TABLE contadorar (
  codmail int(11) NOT NULL default '0'
) TYPE=InnoDB;



#
# Table structure for table 'datoshco'
#

CREATE TABLE datoshco (
  codequipo tinyint(3) unsigned default '0',
  PATH varchar(155) default '0',
  PermisosCambiados tinyint(3) unsigned default '0'
) TYPE=MyISAM COMMENT='Datos historicos';



#
# Table structure for table 'descniveles'
#

CREATE TABLE descniveles (
  codnivel tinyint(3) unsigned NOT NULL default '0',
  descripcion varchar(30) default NULL,
  PRIMARY KEY  (codnivel)
) TYPE=MyISAM;



#
# Table structure for table 'equipos'
#

CREATE TABLE equipos (
  codequipo smallint(3) unsigned NOT NULL default '0',
  descripcion varchar(40) default '0',
  velocidad int(3) unsigned default '0',
  cargaIconsExt tinyint(3) unsigned default '0',
  ExeIntegra varchar(255) default NULL,
  PRIMARY KEY  (codequipo)
) TYPE=MyISAM;



#
# Table structure for table 'extension'
#

CREATE TABLE extension (
  codext tinyint(4) NOT NULL default '0',
  descripcion varchar(30) NOT NULL default '0',
  exten varchar(5) default NULL,
  Modificable tinyint(3) unsigned default NULL,
  Nuevo tinyint(3) unsigned NOT NULL default '0',
  OfertaExe varchar(100) default NULL,
  OfertaPrint varchar(100) default NULL,
  Aparecemenu tinyint(3) unsigned default '0',
  Deshabilitada tinyint(3) unsigned default '0',
  PRIMARY KEY  (codext)
) TYPE=MyISAM;



#
# Table structure for table 'extensionpc'
#

CREATE TABLE extensionpc (
  codext tinyint(3) unsigned NOT NULL default '0',
  codequipo tinyint(3) unsigned NOT NULL default '0',
  pathexe varchar(250) NOT NULL default '',
  impresion varchar(250) default NULL
) TYPE=MyISAM;



#
# Table structure for table 'grupos'
#

CREATE TABLE grupos (
  codgrupo tinyint(3) unsigned NOT NULL default '0',
  nomgrupo varchar(30) NOT NULL default '',
  nivel tinyint(3) unsigned NOT NULL default '0',
  PRIMARY KEY  (codgrupo)
) TYPE=MyISAM;



#
# Table structure for table 'mailc'
#

CREATE TABLE mailc (
  codmail int(10) NOT NULL default '0',
  origen smallint(3) NOT NULL default '0',
  destino smallint(3) NOT NULL default '0',
  leido tinyint(3) NOT NULL default '0',
  email tinyint(3) NOT NULL default '0',
  tipo tinyint(3) NOT NULL default '0',
  PRIMARY KEY  (codmail,origen,destino)
) TYPE=MyISAM COMMENT='Cabecera de mensajeria';



#
# Table structure for table 'mailch'
#

CREATE TABLE mailch (
  codmail int(10) NOT NULL default '0',
  origen smallint(3) NOT NULL default '0',
  destino smallint(3) NOT NULL default '0',
  leido tinyint(3) NOT NULL default '0',
  email tinyint(3) NOT NULL default '0',
  tipo tinyint(3) NOT NULL default '0',
  PRIMARY KEY  (codmail,origen,destino)
) TYPE=MyISAM COMMENT='Cabecera de mensajeria';



#
# Table structure for table 'maildestext'
#

CREATE TABLE maildestext (
  codmail int(10) NOT NULL default '0',
  nombre varchar(255) NOT NULL default '0',
  mail varchar(255) NOT NULL default '0'
) TYPE=MyISAM COMMENT='Destinatarios externos';



#
# Table structure for table 'maildestexth'
#

CREATE TABLE maildestexth (
  codmail int(10) NOT NULL default '0',
  nombre varchar(255) NOT NULL default '0',
  mail varchar(255) NOT NULL default '0'
) TYPE=MyISAM COMMENT='Destinatarios externos';



#
# Table structure for table 'maile'
#

CREATE TABLE maile (
  codmail int(10) NOT NULL default '0',
  origen smallint(3) NOT NULL default '0',
  email tinyint(3) NOT NULL default '0',
  tipo tinyint(3) NOT NULL default '0',
  textoPara varchar(255) NOT NULL default '',
  Destinatarios varchar(255) NOT NULL default '',
  PRIMARY KEY  (codmail,origen)
) TYPE=MyISAM COMMENT='Mensajes enviados';



#
# Table structure for table 'maileh'
#

CREATE TABLE maileh (
  codmail int(10) NOT NULL default '0',
  origen smallint(3) NOT NULL default '0',
  email tinyint(3) NOT NULL default '0',
  tipo tinyint(3) NOT NULL default '0',
  textoPara varchar(255) NOT NULL default '',
  Destinatarios varchar(255) NOT NULL default '',
  PRIMARY KEY  (codmail,origen)
) TYPE=MyISAM COMMENT='Mensajes enviados';



#
# Table structure for table 'maill'
#

CREATE TABLE maill (
  codmail int(3) NOT NULL default '0',
  asunto varchar(200) default '0',
  Texto text,
  Fecha date NOT NULL default '0000-00-00',
  PRIMARY KEY  (codmail)
) TYPE=MyISAM COMMENT='lineas mensajeria';



#
# Table structure for table 'mailtipo'
#

CREATE TABLE mailtipo (
  tipo tinyint(3) NOT NULL default '0',
  Descripcion varchar(35) NOT NULL default '',
  color varchar(15) default NULL,
  numico tinyint(3) unsigned default '0',
  PRIMARY KEY  (tipo)
) TYPE=MyISAM COMMENT='Tipos mensaje';



#
# Table structure for table 'plantilla'
#

CREATE TABLE plantilla (
  codigo tinyint(3) unsigned NOT NULL default '0',
  Descripcion varchar(250) default NULL,
  tipo tinyint(3) unsigned default '0',
  lectura int(11) default '0',
  fecha date default NULL,
  carpeta smallint(5) unsigned NOT NULL default '0',
  PRIMARY KEY  (codigo)
) TYPE=MyISAM COMMENT='Plantillas';



#
# Table structure for table 'preferenciapersonal'
#

CREATE TABLE preferenciapersonal (
  codusu int(11) NOT NULL default '0',
  c1 smallint(6) default '0',
  c2 smallint(6) default '0',
  c3 smallint(6) default '0',
  c4 smallint(6) default '0',
  f1 smallint(6) default '0',
  f2 smallint(6) default '0',
  f3 smallint(6) default '0',
  imp1 smallint(6) default '0',
  imp2 smallint(6) default '0',
  obs smallint(6) default '0',
  tamayo smallint(6) default '0',
  vista tinyint(3) unsigned default '0',
  ancho tinyint(3) unsigned NOT NULL default '20',
  ORDERBY varchar(30) default NULL,
  mailInicio tinyint(3) unsigned default '0',
  mailFiltro tinyint(3) unsigned default '0',
  mailPasarHCO tinyint(3) unsigned default '0',
  PRIMARY KEY  (codusu)
) TYPE=MyISAM;



#
# Table structure for table 'procesos'
#

CREATE TABLE procesos (
  codusu tinyint(3) unsigned NOT NULL default '0',
  codequipo tinyint(3) unsigned NOT NULL default '0',
  proceso int(10) unsigned default NULL,
  fichero varchar(255) NOT NULL default ''
) TYPE=MyISAM;



#
# Table structure for table 'timagen'
#

CREATE TABLE timagen (
  codigo int(10) unsigned NOT NULL default '0',
  codext tinyint(3) unsigned NOT NULL default '0',
  codcarpeta smallint(3) unsigned NOT NULL default '0',
  campo1 varchar(50) NOT NULL default '0',
  campo2 varchar(50) default '0',
  campo3 varchar(50) default '0',
  campo4 varchar(50) default NULL,
  fecha1 date default NULL,
  fecha2 date default NULL,
  fecha3 date default NULL,
  importe1 decimal(14,2) default NULL,
  importe2 decimal(14,2) default NULL,
  observa text,
  tamnyo decimal(10,3) NOT NULL default '0.000',
  userprop int(10) unsigned NOT NULL default '0',
  groupprop int(10) unsigned NOT NULL default '0',
  lecturau int(10) unsigned NOT NULL default '0',
  lecturag int(10) unsigned NOT NULL default '0',
  escriturau int(10) unsigned NOT NULL default '0',
  escriturag int(10) unsigned NOT NULL default '0',
  bloqueo tinyint(3) unsigned NOT NULL default '0',
  PRIMARY KEY  (codigo),
  KEY NewIndex (codcarpeta)
) TYPE=MyISAM COMMENT='Datos archvos';



#
# Table structure for table 'timagenhco'
#

CREATE TABLE timagenhco (
  codequipo tinyint(3) unsigned NOT NULL default '0',
  codigo int(10) unsigned NOT NULL default '0',
  codext tinyint(3) unsigned NOT NULL default '0',
  codcarpeta tinyint(3) unsigned NOT NULL default '0',
  campo1 varchar(50) NOT NULL default '0',
  campo2 varchar(50) default '0',
  campo3 varchar(50) default '0',
  campo4 varchar(50) default NULL,
  fecha1 date default NULL,
  fecha2 date default NULL,
  fecha3 date default NULL,
  importe1 decimal(14,2) default NULL,
  importe2 decimal(14,2) default NULL,
  observa text,
  tamnyo decimal(10,3) NOT NULL default '0.000',
  userprop int(10) unsigned NOT NULL default '0',
  groupprop int(10) unsigned NOT NULL default '0',
  lecturau int(10) unsigned NOT NULL default '0',
  lecturag int(10) unsigned NOT NULL default '0',
  escriturau int(10) unsigned NOT NULL default '0',
  escriturag int(10) unsigned NOT NULL default '0',
  bloqueo tinyint(3) unsigned NOT NULL default '0',
  PRIMARY KEY  (codequipo,codigo),
  KEY NewIndex (codcarpeta)
) TYPE=MyISAM COMMENT='Datos archvos';



#
# Table structure for table 'tmpbusqueda'
#

CREATE TABLE tmpbusqueda (
  codusu tinyint(3) unsigned NOT NULL default '0',
  codequipo tinyint(3) unsigned NOT NULL default '0',
  imagen int(10) unsigned default NULL,
  codcarpeta smallint(5) unsigned NOT NULL default '0'
) TYPE=MyISAM;



#
# Table structure for table 'tmpfich'
#

CREATE TABLE tmpfich (
  codusu tinyint(3) unsigned NOT NULL default '0',
  codequipo tinyint(3) unsigned NOT NULL default '0',
  imagen int(10) unsigned default NULL
) TYPE=MyISAM;



#
# Table structure for table 'tmpintegra'
#

CREATE TABLE tmpintegra (
  codusu tinyint(3) unsigned NOT NULL default '0',
  codigo int(10) unsigned NOT NULL default '0',
  carpeta varchar(100) NOT NULL default '',
  campo1 varchar(50) NOT NULL default '0',
  campo2 varchar(50) default NULL,
  campo3 varchar(50) default NULL,
  campo4 varchar(50) default NULL,
  fecha1 date default NULL,
  fecha2 date default NULL,
  fecha3 date default NULL,
  importe1 decimal(14,2) default NULL,
  importe2 decimal(14,2) default NULL,
  observa text,
  NombreArchivo varchar(100) NOT NULL default '',
  correcto tinyint(3) unsigned NOT NULL default '0',
  codcarpeta smallint(5) unsigned NOT NULL default '0',
  tamanyo int(11) NOT NULL default '0',
  PRIMARY KEY  (codigo,codusu)
) TYPE=MyISAM COMMENT='Integraciones ariadna';



#
# Table structure for table 'usuarios'
#

CREATE TABLE usuarios (
  codusu tinyint(3) unsigned NOT NULL default '0',
  Nombre varchar(40) default '0',
  login varchar(15) default '0',
  password varchar(15) default '0',
  preferencias varchar(50) default NULL,
  maildir varchar(100) default NULL,
  mailserver varchar(100) default NULL,
  mailuser varchar(100) default NULL,
  mailpwd varchar(100) default NULL,
  PRIMARY KEY  (codusu)
) TYPE=MyISAM;



#
# Table structure for table 'usuariosgrupos'
#

CREATE TABLE usuariosgrupos (
  codusu tinyint(3) unsigned NOT NULL default '0',
  codgrupo tinyint(3) unsigned NOT NULL default '0',
  orden tinyint(3) unsigned default NULL,
  PRIMARY KEY  (codgrupo,codusu)
) TYPE=MyISAM COMMENT='Usuarios-grupos';


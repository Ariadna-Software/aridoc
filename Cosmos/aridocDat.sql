# Dumping data for table 'almacen'

INSERT INTO almacen VALUES("0", "1", "/u/aridoc/iconos", "192.168.4.2", "root", "aritel");
INSERT INTO almacen VALUES("3", "1", "/u/aridoc/almacen1", "192.168.4.2", "root", "aritel");
INSERT INTO almacen VALUES("1", "1", "/u/aridoc/files", "192.168.4.2", "root", "aritel");
INSERT INTO almacen VALUES("2", "1", "/u/aridoc/plantillas", "192.168.4.2", "root", "aritel");

# Dumping data for table 'configuracion'

INSERT INTO configuracion VALUES("1", "Nombre", "Auxiliar", "Procedencia", "Otros", "Fecha Doc", "Fecha2", "Fecha3", "Debe", "Haber", "Observaciones", "1");

# Dumping data for table 'descniveles'

INSERT INTO descniveles VALUES("0", "Admin. SISTEMA");
INSERT INTO descniveles VALUES("1", "Administrador");
INSERT INTO descniveles VALUES("3", "Avanzado");
INSERT INTO descniveles VALUES("5", "Normal");
INSERT INTO descniveles VALUES("7", "Bajo");
INSERT INTO descniveles VALUES("10", "Consulta");

# Dumping data for table 'grupos'

INSERT INTO grupos VALUES("1", "Administradores Aridoc", "0");

# Dumping data for table 'mailtipo'

INSERT INTO mailtipo VALUES("0", "GENERICO", " 0", "0");


#Usuario ROOT
INSERT INTO preferenciapersonal (codusu, c1, c2, c3, c4, f1, f2, f3, imp1, imp2, obs, tamayo, vista, ancho, ORDERBY, mailInicio, mailFiltro, mailPasarHCO) VALUES (0, 5000, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 20, NULL, 0, 0, 0);

INSERT INTO usuarios (codusu, Nombre, login, password, preferencias, maildir, mailserver, mailuser, mailpwd) VALUES (0, 'Administrador ARIDOC', 'root', 'aritel', NULL, NULL, NULL, NULL, NULL);

INSERT INTO usuariosgrupos (codusu, codgrupo, orden) VALUES (0, 1, 1);
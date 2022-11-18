CREATE TABLE [Empleado] (
	[Id]	INTEGER NOT NULL CHECK(Id > 0) UNIQUE,
	[ClvEmpleado]	INTEGER NOT NULL CHECK(ClvEmpleado > 0) UNIQUE,
	[Nombre]	TEXT NOT NULL DEFAULT "Sin nombre",
	[ApellidoPaterno]	TEXT NOT NULL DEFAULT "Sin apellido",
	[ApellidoMaterno]	TEXT NOT NULL DEFAULT "Sin apellido",
	[Direccion]	TEXT NOT NULL DEFAULT "Sin direccion",
	[Cargo]	TEXT NOT NULL DEFAULT "Sin cargo",
	[Rfc]	TEXT NOT NULL DEFAULT "Sin RFC",
	[Curp]	TEXT NOT NULL DEFAULT "Sin CURP",
	[Telefono]	TEXT NOT NULL DEFAULT "Sin telefono",
	[HorasSemana]	INTEGER NOT NULL DEFAULT "Sin horas" CHECK(HorasSemana > 0),
	PRIMARY KEY([Id] AUTOINCREMENT)
);



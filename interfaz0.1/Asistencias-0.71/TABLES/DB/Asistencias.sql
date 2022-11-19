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


CREATE UNIQUE INDEX [IX_Empleado_ClvEmpleado] ON [Empleado](ClvEmpleado);


CREATE TABLE [Asistencia] (
	[Id]	INTEGER NOT NULL UNIQUE,
	[Fecha]	TEXT NOT NULL,
	[Llegada]	TEXT NOT NULL,
	[Salida]	TEXT NOT NULL DEFAULT '',
	PRIMARY KEY([Id] AUTOINCREMENT)
);

CREATE INDEX [IX_Asistencia_Fecha] ON [Asistencia]([Fecha]);

ALTER TABLE [Asistencia] ADD COLUMN [ClvEmpleado] INTEGER NOT NULL REFERENCES [Empleado]([ClvEmpleado]);

CREATE TABLE "Admin" (
	"Id"	INTEGER NOT NULL UNIQUE,
	"Password"	TEXT NOT NULL CHECK(length(Password) > 8),
	PRIMARY KEY("Id" AUTOINCREMENT)
);

ALTER TABLE [Admin] ADD COLUMN [ClvEmpleado] INTEGER NOT NULL REFERENCES [Empleado]([ClvEmpleado]);


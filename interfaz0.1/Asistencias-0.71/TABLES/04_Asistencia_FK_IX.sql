CREATE INDEX [IX_Asistencia_Fecha] ON [Asistencia]([Fecha]);

ALTER TABLE [Asistencia] ADD COLUMN [ClvEmpleado] INTEGER NOT NULL REFERENCES [Empleado]([ClvEmpleado]);

-- Portafolio QA: esquema idempotente corregido. Respaldar y seleccionar base en SSMS.
SET ANSI_NULLS ON;
SET QUOTED_IDENTIFIER ON;
SET ANSI_PADDING ON;
SET ANSI_WARNINGS ON;
SET CONCAT_NULL_YIELDS_NULL ON;
SET ARITHABORT ON;
SET NUMERIC_ROUNDABORT OFF;
GO
IF OBJECT_ID(N'[__EFMigrationsHistory]') IS NULL
BEGIN
    CREATE TABLE [__EFMigrationsHistory] (
        [MigrationId] nvarchar(150) NOT NULL,
        [ProductVersion] nvarchar(32) NOT NULL,
        CONSTRAINT [PK___EFMigrationsHistory] PRIMARY KEY ([MigrationId])
    );
END;
GO

BEGIN TRANSACTION;
GO

IF NOT EXISTS (
    SELECT * FROM [__EFMigrationsHistory]
    WHERE [MigrationId] = N'20261007003147_InicialEtapa1'
)
BEGIN
    CREATE TABLE [Instituciones] (
        [Id] int NOT NULL IDENTITY,
        [Nombre] nvarchar(150) NOT NULL,
        [Activo] bit NOT NULL,
        CONSTRAINT [PK_Instituciones] PRIMARY KEY ([Id])
    );
END;
GO

IF NOT EXISTS (
    SELECT * FROM [__EFMigrationsHistory]
    WHERE [MigrationId] = N'20261007003147_InicialEtapa1'
)
BEGIN
    CREATE TABLE [RecursosQA] (
        [Id] int NOT NULL IDENTITY,
        [Nombres] nvarchar(100) NOT NULL,
        [Apellidos] nvarchar(100) NOT NULL,
        [Correo] nvarchar(150) NOT NULL,
        [Equipo] int NOT NULL,
        [CapacidadDisponible] decimal(5,2) NOT NULL,
        [Activo] bit NOT NULL,
        CONSTRAINT [PK_RecursosQA] PRIMARY KEY ([Id])
    );
END;
GO

IF NOT EXISTS (
    SELECT * FROM [__EFMigrationsHistory]
    WHERE [MigrationId] = N'20261007003147_InicialEtapa1'
)
BEGIN
    CREATE TABLE [Sistemas] (
        [Id] int NOT NULL IDENTITY,
        [Nombre] nvarchar(150) NOT NULL,
        [Activo] bit NOT NULL,
        CONSTRAINT [PK_Sistemas] PRIMARY KEY ([Id])
    );
END;
GO

IF NOT EXISTS (
    SELECT * FROM [__EFMigrationsHistory]
    WHERE [MigrationId] = N'20261007003147_InicialEtapa1'
)
BEGIN
    CREATE TABLE [Proyectos] (
        [Id] int NOT NULL IDENTITY,
        [Nombre] nvarchar(200) NOT NULL,
        [Descripcion] nvarchar(500) NULL,
        [InstitucionId] int NOT NULL,
        [Tipo] int NOT NULL,
        [Estado] int NOT NULL,
        [PM] nvarchar(150) NULL,
        [Lider] nvarchar(150) NULL,
        [FechaInicio] datetime2 NOT NULL,
        [FechaFin] datetime2 NOT NULL,
        [Avance] decimal(5,2) NOT NULL,
        [Activo] bit NOT NULL,
        CONSTRAINT [PK_Proyectos] PRIMARY KEY ([Id]),
        CONSTRAINT [FK_Proyectos_Instituciones_InstitucionId] FOREIGN KEY ([InstitucionId]) REFERENCES [Instituciones] ([Id]) ON DELETE NO ACTION
    );
END;
GO

IF NOT EXISTS (
    SELECT * FROM [__EFMigrationsHistory]
    WHERE [MigrationId] = N'20261007003147_InicialEtapa1'
)
BEGIN
    CREATE TABLE [SistemaInstituciones] (
        [SistemaId] int NOT NULL,
        [InstitucionId] int NOT NULL,
        CONSTRAINT [PK_SistemaInstituciones] PRIMARY KEY ([SistemaId], [InstitucionId]),
        CONSTRAINT [FK_SistemaInstituciones_Instituciones_InstitucionId] FOREIGN KEY ([InstitucionId]) REFERENCES [Instituciones] ([Id]) ON DELETE CASCADE,
        CONSTRAINT [FK_SistemaInstituciones_Sistemas_SistemaId] FOREIGN KEY ([SistemaId]) REFERENCES [Sistemas] ([Id]) ON DELETE CASCADE
    );
END;
GO

IF NOT EXISTS (
    SELECT * FROM [__EFMigrationsHistory]
    WHERE [MigrationId] = N'20261007003147_InicialEtapa1'
)
BEGIN
    CREATE TABLE [Asignaciones] (
        [Id] int NOT NULL IDENTITY,
        [ProyectoId] int NOT NULL,
        [RecursoQAId] int NOT NULL,
        [FechaInicio] datetime2 NOT NULL,
        [FechaFin] datetime2 NOT NULL,
        [Porcentaje] decimal(5,2) NOT NULL,
        [Rol] nvarchar(100) NOT NULL,
        CONSTRAINT [PK_Asignaciones] PRIMARY KEY ([Id]),
        CONSTRAINT [FK_Asignaciones_Proyectos_ProyectoId] FOREIGN KEY ([ProyectoId]) REFERENCES [Proyectos] ([Id]) ON DELETE CASCADE,
        CONSTRAINT [FK_Asignaciones_RecursosQA_RecursoQAId] FOREIGN KEY ([RecursoQAId]) REFERENCES [RecursosQA] ([Id]) ON DELETE NO ACTION
    );
END;
GO

IF NOT EXISTS (
    SELECT * FROM [__EFMigrationsHistory]
    WHERE [MigrationId] = N'20261007003147_InicialEtapa1'
)
BEGIN
    CREATE TABLE [AsociacionesExternas] (
        [Id] int NOT NULL IDENTITY,
        [ProyectoId] int NOT NULL,
        [Herramienta] int NOT NULL,
        [IdentificadorExterno] nvarchar(150) NOT NULL,
        [Estado] int NOT NULL,
        [UltimaSincronizacion] datetime2 NULL,
        [UltimoError] nvarchar(500) NULL,
        CONSTRAINT [PK_AsociacionesExternas] PRIMARY KEY ([Id]),
        CONSTRAINT [FK_AsociacionesExternas_Proyectos_ProyectoId] FOREIGN KEY ([ProyectoId]) REFERENCES [Proyectos] ([Id]) ON DELETE CASCADE
    );
END;
GO

IF NOT EXISTS (
    SELECT * FROM [__EFMigrationsHistory]
    WHERE [MigrationId] = N'20261007003147_InicialEtapa1'
)
BEGIN
    CREATE TABLE [ProyectoSistemas] (
        [ProyectoId] int NOT NULL,
        [SistemaId] int NOT NULL,
        CONSTRAINT [PK_ProyectoSistemas] PRIMARY KEY ([ProyectoId], [SistemaId]),
        CONSTRAINT [FK_ProyectoSistemas_Proyectos_ProyectoId] FOREIGN KEY ([ProyectoId]) REFERENCES [Proyectos] ([Id]) ON DELETE CASCADE,
        CONSTRAINT [FK_ProyectoSistemas_Sistemas_SistemaId] FOREIGN KEY ([SistemaId]) REFERENCES [Sistemas] ([Id]) ON DELETE CASCADE
    );
END;
GO

IF NOT EXISTS (
    SELECT * FROM [__EFMigrationsHistory]
    WHERE [MigrationId] = N'20261007003147_InicialEtapa1'
)
BEGIN
    CREATE INDEX [IX_Asignaciones_ProyectoId] ON [Asignaciones] ([ProyectoId]);
END;
GO

IF NOT EXISTS (
    SELECT * FROM [__EFMigrationsHistory]
    WHERE [MigrationId] = N'20261007003147_InicialEtapa1'
)
BEGIN
    CREATE INDEX [IX_Asignaciones_RecursoQAId] ON [Asignaciones] ([RecursoQAId]);
END;
GO

IF NOT EXISTS (
    SELECT * FROM [__EFMigrationsHistory]
    WHERE [MigrationId] = N'20261007003147_InicialEtapa1'
)
BEGIN
    CREATE INDEX [IX_AsociacionesExternas_ProyectoId] ON [AsociacionesExternas] ([ProyectoId]);
END;
GO

IF NOT EXISTS (
    SELECT * FROM [__EFMigrationsHistory]
    WHERE [MigrationId] = N'20261007003147_InicialEtapa1'
)
BEGIN
    CREATE UNIQUE INDEX [IX_Instituciones_Nombre] ON [Instituciones] ([Nombre]);
END;
GO

IF NOT EXISTS (
    SELECT * FROM [__EFMigrationsHistory]
    WHERE [MigrationId] = N'20261007003147_InicialEtapa1'
)
BEGIN
    CREATE INDEX [IX_Proyectos_InstitucionId] ON [Proyectos] ([InstitucionId]);
END;
GO

IF NOT EXISTS (
    SELECT * FROM [__EFMigrationsHistory]
    WHERE [MigrationId] = N'20261007003147_InicialEtapa1'
)
BEGIN
    CREATE INDEX [IX_ProyectoSistemas_SistemaId] ON [ProyectoSistemas] ([SistemaId]);
END;
GO

IF NOT EXISTS (
    SELECT * FROM [__EFMigrationsHistory]
    WHERE [MigrationId] = N'20261007003147_InicialEtapa1'
)
BEGIN
    CREATE UNIQUE INDEX [IX_RecursosQA_Correo] ON [RecursosQA] ([Correo]);
END;
GO

IF NOT EXISTS (
    SELECT * FROM [__EFMigrationsHistory]
    WHERE [MigrationId] = N'20261007003147_InicialEtapa1'
)
BEGIN
    CREATE INDEX [IX_SistemaInstituciones_InstitucionId] ON [SistemaInstituciones] ([InstitucionId]);
END;
GO

IF NOT EXISTS (
    SELECT * FROM [__EFMigrationsHistory]
    WHERE [MigrationId] = N'20261007003147_InicialEtapa1'
)
BEGIN
    CREATE UNIQUE INDEX [IX_Sistemas_Nombre] ON [Sistemas] ([Nombre]);
END;
GO

IF NOT EXISTS (
    SELECT * FROM [__EFMigrationsHistory]
    WHERE [MigrationId] = N'20261007003147_InicialEtapa1'
)
BEGIN
    INSERT INTO [__EFMigrationsHistory] ([MigrationId], [ProductVersion])
    VALUES (N'20261007003147_InicialEtapa1', N'8.0.11');
END;
GO

COMMIT;
GO

BEGIN TRANSACTION;
GO

IF NOT EXISTS (
    SELECT * FROM [__EFMigrationsHistory]
    WHERE [MigrationId] = N'20261007154154_Etapa2Planificacion'
)
BEGIN
    CREATE TABLE [Ausencias] (
        [Id] int NOT NULL IDENTITY,
        [RecursoQAId] int NOT NULL,
        [FechaInicio] datetime2 NOT NULL,
        [FechaFin] datetime2 NOT NULL,
        [Motivo] nvarchar(150) NULL,
        CONSTRAINT [PK_Ausencias] PRIMARY KEY ([Id]),
        CONSTRAINT [FK_Ausencias_RecursosQA_RecursoQAId] FOREIGN KEY ([RecursoQAId]) REFERENCES [RecursosQA] ([Id]) ON DELETE NO ACTION
    );
END;
GO

IF NOT EXISTS (
    SELECT * FROM [__EFMigrationsHistory]
    WHERE [MigrationId] = N'20261007154154_Etapa2Planificacion'
)
BEGIN
    CREATE TABLE [Feriados] (
        [Id] int NOT NULL IDENTITY,
        [Fecha] datetime2 NOT NULL,
        [Nombre] nvarchar(150) NOT NULL,
        CONSTRAINT [PK_Feriados] PRIMARY KEY ([Id])
    );
END;
GO

IF NOT EXISTS (
    SELECT * FROM [__EFMigrationsHistory]
    WHERE [MigrationId] = N'20261007154154_Etapa2Planificacion'
)
BEGIN
    CREATE TABLE [Hitos] (
        [Id] int NOT NULL IDENTITY,
        [ProyectoId] int NOT NULL,
        [Fecha] datetime2 NOT NULL,
        [Nombre] nvarchar(150) NOT NULL,
        CONSTRAINT [PK_Hitos] PRIMARY KEY ([Id]),
        CONSTRAINT [FK_Hitos_Proyectos_ProyectoId] FOREIGN KEY ([ProyectoId]) REFERENCES [Proyectos] ([Id]) ON DELETE CASCADE
    );
END;
GO

IF NOT EXISTS (
    SELECT * FROM [__EFMigrationsHistory]
    WHERE [MigrationId] = N'20261007154154_Etapa2Planificacion'
)
BEGIN
    CREATE INDEX [IX_Ausencias_RecursoQAId] ON [Ausencias] ([RecursoQAId]);
END;
GO

IF NOT EXISTS (
    SELECT * FROM [__EFMigrationsHistory]
    WHERE [MigrationId] = N'20261007154154_Etapa2Planificacion'
)
BEGIN
    CREATE UNIQUE INDEX [IX_Feriados_Fecha] ON [Feriados] ([Fecha]);
END;
GO

IF NOT EXISTS (
    SELECT * FROM [__EFMigrationsHistory]
    WHERE [MigrationId] = N'20261007154154_Etapa2Planificacion'
)
BEGIN
    CREATE INDEX [IX_Hitos_ProyectoId] ON [Hitos] ([ProyectoId]);
END;
GO

IF NOT EXISTS (
    SELECT * FROM [__EFMigrationsHistory]
    WHERE [MigrationId] = N'20261007154154_Etapa2Planificacion'
)
BEGIN
    INSERT INTO [__EFMigrationsHistory] ([MigrationId], [ProductVersion])
    VALUES (N'20261007154154_Etapa2Planificacion', N'8.0.11');
END;
GO

COMMIT;
GO

BEGIN TRANSACTION;
GO

IF NOT EXISTS (
    SELECT * FROM [__EFMigrationsHistory]
    WHERE [MigrationId] = N'20261007160421_Etapa3Seguridad'
)
BEGIN
    CREATE TABLE [AspNetRoles] (
        [Id] nvarchar(450) NOT NULL,
        [Name] nvarchar(256) NULL,
        [NormalizedName] nvarchar(256) NULL,
        [ConcurrencyStamp] nvarchar(max) NULL,
        CONSTRAINT [PK_AspNetRoles] PRIMARY KEY ([Id])
    );
END;
GO

IF NOT EXISTS (
    SELECT * FROM [__EFMigrationsHistory]
    WHERE [MigrationId] = N'20261007160421_Etapa3Seguridad'
)
BEGIN
    CREATE TABLE [AspNetUsers] (
        [Id] nvarchar(450) NOT NULL,
        [NombreCompleto] nvarchar(max) NOT NULL,
        [Activo] bit NOT NULL,
        [UserName] nvarchar(256) NULL,
        [NormalizedUserName] nvarchar(256) NULL,
        [Email] nvarchar(256) NULL,
        [NormalizedEmail] nvarchar(256) NULL,
        [EmailConfirmed] bit NOT NULL,
        [PasswordHash] nvarchar(max) NULL,
        [SecurityStamp] nvarchar(max) NULL,
        [ConcurrencyStamp] nvarchar(max) NULL,
        [PhoneNumber] nvarchar(max) NULL,
        [PhoneNumberConfirmed] bit NOT NULL,
        [TwoFactorEnabled] bit NOT NULL,
        [LockoutEnd] datetimeoffset NULL,
        [LockoutEnabled] bit NOT NULL,
        [AccessFailedCount] int NOT NULL,
        CONSTRAINT [PK_AspNetUsers] PRIMARY KEY ([Id])
    );
END;
GO

IF NOT EXISTS (
    SELECT * FROM [__EFMigrationsHistory]
    WHERE [MigrationId] = N'20261007160421_Etapa3Seguridad'
)
BEGIN
    CREATE TABLE [AspNetRoleClaims] (
        [Id] int NOT NULL IDENTITY,
        [RoleId] nvarchar(450) NOT NULL,
        [ClaimType] nvarchar(max) NULL,
        [ClaimValue] nvarchar(max) NULL,
        CONSTRAINT [PK_AspNetRoleClaims] PRIMARY KEY ([Id]),
        CONSTRAINT [FK_AspNetRoleClaims_AspNetRoles_RoleId] FOREIGN KEY ([RoleId]) REFERENCES [AspNetRoles] ([Id]) ON DELETE CASCADE
    );
END;
GO

IF NOT EXISTS (
    SELECT * FROM [__EFMigrationsHistory]
    WHERE [MigrationId] = N'20261007160421_Etapa3Seguridad'
)
BEGIN
    CREATE TABLE [AspNetUserClaims] (
        [Id] int NOT NULL IDENTITY,
        [UserId] nvarchar(450) NOT NULL,
        [ClaimType] nvarchar(max) NULL,
        [ClaimValue] nvarchar(max) NULL,
        CONSTRAINT [PK_AspNetUserClaims] PRIMARY KEY ([Id]),
        CONSTRAINT [FK_AspNetUserClaims_AspNetUsers_UserId] FOREIGN KEY ([UserId]) REFERENCES [AspNetUsers] ([Id]) ON DELETE CASCADE
    );
END;
GO

IF NOT EXISTS (
    SELECT * FROM [__EFMigrationsHistory]
    WHERE [MigrationId] = N'20261007160421_Etapa3Seguridad'
)
BEGIN
    CREATE TABLE [AspNetUserLogins] (
        [LoginProvider] nvarchar(450) NOT NULL,
        [ProviderKey] nvarchar(450) NOT NULL,
        [ProviderDisplayName] nvarchar(max) NULL,
        [UserId] nvarchar(450) NOT NULL,
        CONSTRAINT [PK_AspNetUserLogins] PRIMARY KEY ([LoginProvider], [ProviderKey]),
        CONSTRAINT [FK_AspNetUserLogins_AspNetUsers_UserId] FOREIGN KEY ([UserId]) REFERENCES [AspNetUsers] ([Id]) ON DELETE CASCADE
    );
END;
GO

IF NOT EXISTS (
    SELECT * FROM [__EFMigrationsHistory]
    WHERE [MigrationId] = N'20261007160421_Etapa3Seguridad'
)
BEGIN
    CREATE TABLE [AspNetUserRoles] (
        [UserId] nvarchar(450) NOT NULL,
        [RoleId] nvarchar(450) NOT NULL,
        CONSTRAINT [PK_AspNetUserRoles] PRIMARY KEY ([UserId], [RoleId]),
        CONSTRAINT [FK_AspNetUserRoles_AspNetRoles_RoleId] FOREIGN KEY ([RoleId]) REFERENCES [AspNetRoles] ([Id]) ON DELETE CASCADE,
        CONSTRAINT [FK_AspNetUserRoles_AspNetUsers_UserId] FOREIGN KEY ([UserId]) REFERENCES [AspNetUsers] ([Id]) ON DELETE CASCADE
    );
END;
GO

IF NOT EXISTS (
    SELECT * FROM [__EFMigrationsHistory]
    WHERE [MigrationId] = N'20261007160421_Etapa3Seguridad'
)
BEGIN
    CREATE TABLE [AspNetUserTokens] (
        [UserId] nvarchar(450) NOT NULL,
        [LoginProvider] nvarchar(450) NOT NULL,
        [Name] nvarchar(450) NOT NULL,
        [Value] nvarchar(max) NULL,
        CONSTRAINT [PK_AspNetUserTokens] PRIMARY KEY ([UserId], [LoginProvider], [Name]),
        CONSTRAINT [FK_AspNetUserTokens_AspNetUsers_UserId] FOREIGN KEY ([UserId]) REFERENCES [AspNetUsers] ([Id]) ON DELETE CASCADE
    );
END;
GO

IF NOT EXISTS (
    SELECT * FROM [__EFMigrationsHistory]
    WHERE [MigrationId] = N'20261007160421_Etapa3Seguridad'
)
BEGIN
    CREATE INDEX [IX_AspNetRoleClaims_RoleId] ON [AspNetRoleClaims] ([RoleId]);
END;
GO

IF NOT EXISTS (
    SELECT * FROM [__EFMigrationsHistory]
    WHERE [MigrationId] = N'20261007160421_Etapa3Seguridad'
)
BEGIN
    EXEC(N'CREATE UNIQUE INDEX [RoleNameIndex] ON [AspNetRoles] ([NormalizedName]) WHERE [NormalizedName] IS NOT NULL');
END;
GO

IF NOT EXISTS (
    SELECT * FROM [__EFMigrationsHistory]
    WHERE [MigrationId] = N'20261007160421_Etapa3Seguridad'
)
BEGIN
    CREATE INDEX [IX_AspNetUserClaims_UserId] ON [AspNetUserClaims] ([UserId]);
END;
GO

IF NOT EXISTS (
    SELECT * FROM [__EFMigrationsHistory]
    WHERE [MigrationId] = N'20261007160421_Etapa3Seguridad'
)
BEGIN
    CREATE INDEX [IX_AspNetUserLogins_UserId] ON [AspNetUserLogins] ([UserId]);
END;
GO

IF NOT EXISTS (
    SELECT * FROM [__EFMigrationsHistory]
    WHERE [MigrationId] = N'20261007160421_Etapa3Seguridad'
)
BEGIN
    CREATE INDEX [IX_AspNetUserRoles_RoleId] ON [AspNetUserRoles] ([RoleId]);
END;
GO

IF NOT EXISTS (
    SELECT * FROM [__EFMigrationsHistory]
    WHERE [MigrationId] = N'20261007160421_Etapa3Seguridad'
)
BEGIN
    CREATE INDEX [EmailIndex] ON [AspNetUsers] ([NormalizedEmail]);
END;
GO

IF NOT EXISTS (
    SELECT * FROM [__EFMigrationsHistory]
    WHERE [MigrationId] = N'20261007160421_Etapa3Seguridad'
)
BEGIN
    EXEC(N'CREATE UNIQUE INDEX [UserNameIndex] ON [AspNetUsers] ([NormalizedUserName]) WHERE [NormalizedUserName] IS NOT NULL');
END;
GO

IF NOT EXISTS (
    SELECT * FROM [__EFMigrationsHistory]
    WHERE [MigrationId] = N'20261007160421_Etapa3Seguridad'
)
BEGIN
    INSERT INTO [__EFMigrationsHistory] ([MigrationId], [ProductVersion])
    VALUES (N'20261007160421_Etapa3Seguridad', N'8.0.11');
END;
GO

COMMIT;
GO

BEGIN TRANSACTION;
GO

IF NOT EXISTS (
    SELECT * FROM [__EFMigrationsHistory]
    WHERE [MigrationId] = N'20261007210436_Etapa4Actividades'
)
BEGIN
    CREATE TABLE [Actividades] (
        [Id] int NOT NULL IDENTITY,
        [ProyectoId] int NOT NULL,
        [Nombre] nvarchar(200) NOT NULL,
        [FechaInicio] datetime2 NOT NULL,
        [FechaFin] datetime2 NOT NULL,
        [Estado] int NOT NULL,
        CONSTRAINT [PK_Actividades] PRIMARY KEY ([Id]),
        CONSTRAINT [FK_Actividades_Proyectos_ProyectoId] FOREIGN KEY ([ProyectoId]) REFERENCES [Proyectos] ([Id]) ON DELETE CASCADE
    );
END;
GO

IF NOT EXISTS (
    SELECT * FROM [__EFMigrationsHistory]
    WHERE [MigrationId] = N'20261007210436_Etapa4Actividades'
)
BEGIN
    CREATE INDEX [IX_Actividades_ProyectoId] ON [Actividades] ([ProyectoId]);
END;
GO

IF NOT EXISTS (
    SELECT * FROM [__EFMigrationsHistory]
    WHERE [MigrationId] = N'20261007210436_Etapa4Actividades'
)
BEGIN
    INSERT INTO [__EFMigrationsHistory] ([MigrationId], [ProductVersion])
    VALUES (N'20261007210436_Etapa4Actividades', N'8.0.11');
END;
GO

COMMIT;
GO

BEGIN TRANSACTION;
GO

IF NOT EXISTS (
    SELECT * FROM [__EFMigrationsHistory]
    WHERE [MigrationId] = N'20261007214134_Etapa5AvanceActividad'
)
BEGIN
    ALTER TABLE [Actividades] ADD [Avance] decimal(5,2) NOT NULL DEFAULT 0.0;
END;
GO

IF NOT EXISTS (
    SELECT * FROM [__EFMigrationsHistory]
    WHERE [MigrationId] = N'20261007214134_Etapa5AvanceActividad'
)
BEGIN
    INSERT INTO [__EFMigrationsHistory] ([MigrationId], [ProductVersion])
    VALUES (N'20261007214134_Etapa5AvanceActividad', N'8.0.11');
END;
GO

COMMIT;
GO

BEGIN TRANSACTION;
GO

IF NOT EXISTS (
    SELECT * FROM [__EFMigrationsHistory]
    WHERE [MigrationId] = N'20261008003316_Etapa6UsuariosNombresRecursoVinculado'
)
BEGIN
    ALTER TABLE [RecursosQA] ADD [UsuarioId] nvarchar(450) NULL;
END;
GO

IF NOT EXISTS (
    SELECT * FROM [__EFMigrationsHistory]
    WHERE [MigrationId] = N'20261008003316_Etapa6UsuariosNombresRecursoVinculado'
)
BEGIN
    ALTER TABLE [AspNetUsers] ADD [Apellidos] nvarchar(100) NOT NULL DEFAULT N'';
END;
GO

IF NOT EXISTS (
    SELECT * FROM [__EFMigrationsHistory]
    WHERE [MigrationId] = N'20261008003316_Etapa6UsuariosNombresRecursoVinculado'
)
BEGIN
    ALTER TABLE [AspNetUsers] ADD [Nombres] nvarchar(100) NOT NULL DEFAULT N'';
END;
GO

IF NOT EXISTS (
    SELECT * FROM [__EFMigrationsHistory]
    WHERE [MigrationId] = N'20261008003316_Etapa6UsuariosNombresRecursoVinculado'
)
BEGIN

    EXEC(N'UPDATE AspNetUsers SET
        Nombres   = CASE WHEN CHARINDEX('' '', LTRIM(RTRIM(NombreCompleto))) > 0
                         THEN LEFT(LTRIM(RTRIM(NombreCompleto)), CHARINDEX('' '', LTRIM(RTRIM(NombreCompleto))) - 1)
                         ELSE LTRIM(RTRIM(NombreCompleto)) END,
        Apellidos = CASE WHEN CHARINDEX('' '', LTRIM(RTRIM(NombreCompleto))) > 0
                         THEN LTRIM(SUBSTRING(LTRIM(RTRIM(NombreCompleto)), CHARINDEX('' '', LTRIM(RTRIM(NombreCompleto))) + 1, 200))
                         ELSE '''' END;');
END;
GO

IF NOT EXISTS (
    SELECT * FROM [__EFMigrationsHistory]
    WHERE [MigrationId] = N'20261008003316_Etapa6UsuariosNombresRecursoVinculado'
)
BEGIN
    DECLARE @var0 sysname;
    SELECT @var0 = [d].[name]
    FROM [sys].[default_constraints] [d]
    INNER JOIN [sys].[columns] [c] ON [d].[parent_column_id] = [c].[column_id] AND [d].[parent_object_id] = [c].[object_id]
    WHERE ([d].[parent_object_id] = OBJECT_ID(N'[AspNetUsers]') AND [c].[name] = N'NombreCompleto');
    IF @var0 IS NOT NULL EXEC(N'ALTER TABLE [AspNetUsers] DROP CONSTRAINT [' + @var0 + '];');
    ALTER TABLE [AspNetUsers] DROP COLUMN [NombreCompleto];
END;
GO

IF NOT EXISTS (
    SELECT * FROM [__EFMigrationsHistory]
    WHERE [MigrationId] = N'20261008003316_Etapa6UsuariosNombresRecursoVinculado'
)
BEGIN
    EXEC(N'CREATE UNIQUE INDEX [IX_RecursosQA_UsuarioId] ON [RecursosQA] ([UsuarioId]) WHERE [UsuarioId] IS NOT NULL');
END;
GO

IF NOT EXISTS (
    SELECT * FROM [__EFMigrationsHistory]
    WHERE [MigrationId] = N'20261008003316_Etapa6UsuariosNombresRecursoVinculado'
)
BEGIN
    ALTER TABLE [RecursosQA] ADD CONSTRAINT [FK_RecursosQA_AspNetUsers_UsuarioId] FOREIGN KEY ([UsuarioId]) REFERENCES [AspNetUsers] ([Id]) ON DELETE NO ACTION;
END;
GO

IF NOT EXISTS (
    SELECT * FROM [__EFMigrationsHistory]
    WHERE [MigrationId] = N'20261008003316_Etapa6UsuariosNombresRecursoVinculado'
)
BEGIN

    EXEC(N'UPDATE r SET r.UsuarioId = u.Id
    FROM RecursosQA r
    JOIN AspNetUsers u ON u.NormalizedEmail = UPPER(r.Correo)
    WHERE r.UsuarioId IS NULL;');
END;
GO

IF NOT EXISTS (
    SELECT * FROM [__EFMigrationsHistory]
    WHERE [MigrationId] = N'20261008003316_Etapa6UsuariosNombresRecursoVinculado'
)
BEGIN
    INSERT INTO [__EFMigrationsHistory] ([MigrationId], [ProductVersion])
    VALUES (N'20261008003316_Etapa6UsuariosNombresRecursoVinculado', N'8.0.11');
END;
GO

COMMIT;
GO

BEGIN TRANSACTION;
GO

IF NOT EXISTS (
    SELECT * FROM [__EFMigrationsHistory]
    WHERE [MigrationId] = N'20261008005910_Etapa7RecursoEnActividad'
)
BEGIN
    ALTER TABLE [Actividades] ADD [RecursoQAId] int NULL;
END;
GO

IF NOT EXISTS (
    SELECT * FROM [__EFMigrationsHistory]
    WHERE [MigrationId] = N'20261008005910_Etapa7RecursoEnActividad'
)
BEGIN
    CREATE INDEX [IX_Actividades_RecursoQAId] ON [Actividades] ([RecursoQAId]);
END;
GO

IF NOT EXISTS (
    SELECT * FROM [__EFMigrationsHistory]
    WHERE [MigrationId] = N'20261008005910_Etapa7RecursoEnActividad'
)
BEGIN
    ALTER TABLE [Actividades] ADD CONSTRAINT [FK_Actividades_RecursosQA_RecursoQAId] FOREIGN KEY ([RecursoQAId]) REFERENCES [RecursosQA] ([Id]) ON DELETE NO ACTION;
END;
GO

IF NOT EXISTS (
    SELECT * FROM [__EFMigrationsHistory]
    WHERE [MigrationId] = N'20261008005910_Etapa7RecursoEnActividad'
)
BEGIN
    INSERT INTO [__EFMigrationsHistory] ([MigrationId], [ProductVersion])
    VALUES (N'20261008005910_Etapa7RecursoEnActividad', N'8.0.11');
END;
GO

COMMIT;
GO

BEGIN TRANSACTION;
GO

IF NOT EXISTS (
    SELECT * FROM [__EFMigrationsHistory]
    WHERE [MigrationId] = N'20261009144611_Etapa8EstadoPorAvance'
)
BEGIN

    UPDATE p SET p.Avance = ISNULL((SELECT ROUND(AVG(a.Avance), 2) FROM Actividades a WHERE a.ProyectoId = p.Id AND a.Estado <> 3), 100)
    FROM Proyectos p
    WHERE p.Estado <> 4 AND EXISTS (SELECT 1 FROM Actividades a WHERE a.ProyectoId = p.Id);
END;
GO

IF NOT EXISTS (
    SELECT * FROM [__EFMigrationsHistory]
    WHERE [MigrationId] = N'20261009144611_Etapa8EstadoPorAvance'
)
BEGIN

    UPDATE Proyectos SET Estado = CASE WHEN Avance >= 100 THEN 3 WHEN Avance <= 0 THEN 0 ELSE 2 END
    WHERE Estado IN (0, 1, 2, 3);
END;
GO

IF NOT EXISTS (
    SELECT * FROM [__EFMigrationsHistory]
    WHERE [MigrationId] = N'20261009144611_Etapa8EstadoPorAvance'
)
BEGIN
    INSERT INTO [__EFMigrationsHistory] ([MigrationId], [ProductVersion])
    VALUES (N'20261009144611_Etapa8EstadoPorAvance', N'8.0.11');
END;
GO

COMMIT;
GO

BEGIN TRANSACTION;
GO

IF NOT EXISTS (
    SELECT * FROM [__EFMigrationsHistory]
    WHERE [MigrationId] = N'20261009150116_Etapa9TrazabilidadCierre'
)
BEGIN
    ALTER TABLE [Asignaciones] ADD [FechaFinOriginal] datetime2 NULL;
END;
GO

IF NOT EXISTS (
    SELECT * FROM [__EFMigrationsHistory]
    WHERE [MigrationId] = N'20261009150116_Etapa9TrazabilidadCierre'
)
BEGIN
    ALTER TABLE [Asignaciones] ADD [NoEjecutada] bit NOT NULL DEFAULT CAST(0 AS bit);
END;
GO

IF NOT EXISTS (
    SELECT * FROM [__EFMigrationsHistory]
    WHERE [MigrationId] = N'20261009150116_Etapa9TrazabilidadCierre'
)
BEGIN
    CREATE TABLE [ActividadAuditorias] (
        [Id] int NOT NULL IDENTITY,
        [ProyectoId] int NOT NULL,
        [ActividadId] int NULL,
        [Actividad] nvarchar(200) NOT NULL,
        [Fecha] datetime2 NOT NULL,
        [Usuario] nvarchar(256) NOT NULL,
        [Accion] nvarchar(40) NOT NULL,
        [Detalle] nvarchar(500) NULL,
        [ProyectoCerrado] bit NOT NULL,
        CONSTRAINT [PK_ActividadAuditorias] PRIMARY KEY ([Id]),
        CONSTRAINT [FK_ActividadAuditorias_Proyectos_ProyectoId] FOREIGN KEY ([ProyectoId]) REFERENCES [Proyectos] ([Id]) ON DELETE CASCADE
    );
END;
GO

IF NOT EXISTS (
    SELECT * FROM [__EFMigrationsHistory]
    WHERE [MigrationId] = N'20261009150116_Etapa9TrazabilidadCierre'
)
BEGIN
    CREATE TABLE [Cierres] (
        [Id] int NOT NULL IDENTITY,
        [ProyectoId] int NOT NULL,
        [Fecha] datetime2 NOT NULL,
        [Usuario] nvarchar(256) NOT NULL,
        [AvanceFinal] decimal(5,2) NOT NULL,
        [TotalActividades] int NOT NULL,
        [ActividadesCompletadas] int NOT NULL,
        [Reconstruido] bit NOT NULL,
        [Vigente] bit NOT NULL,
        [FechaReapertura] datetime2 NULL,
        [UsuarioReapertura] nvarchar(256) NULL,
        CONSTRAINT [PK_Cierres] PRIMARY KEY ([Id]),
        CONSTRAINT [FK_Cierres_Proyectos_ProyectoId] FOREIGN KEY ([ProyectoId]) REFERENCES [Proyectos] ([Id]) ON DELETE CASCADE
    );
END;
GO

IF NOT EXISTS (
    SELECT * FROM [__EFMigrationsHistory]
    WHERE [MigrationId] = N'20261009150116_Etapa9TrazabilidadCierre'
)
BEGIN
    CREATE TABLE [CierreActividades] (
        [Id] int NOT NULL IDENTITY,
        [CierreProyectoRegistroId] int NOT NULL,
        [Nombre] nvarchar(200) NOT NULL,
        [Responsable] nvarchar(200) NULL,
        [Estado] nvarchar(30) NOT NULL,
        [Avance] decimal(5,2) NOT NULL,
        [FechaInicio] datetime2 NOT NULL,
        [FechaFin] datetime2 NOT NULL,
        CONSTRAINT [PK_CierreActividades] PRIMARY KEY ([Id]),
        CONSTRAINT [FK_CierreActividades_Cierres_CierreProyectoRegistroId] FOREIGN KEY ([CierreProyectoRegistroId]) REFERENCES [Cierres] ([Id]) ON DELETE CASCADE
    );
END;
GO

IF NOT EXISTS (
    SELECT * FROM [__EFMigrationsHistory]
    WHERE [MigrationId] = N'20261009150116_Etapa9TrazabilidadCierre'
)
BEGIN
    CREATE TABLE [CierreParticipantes] (
        [Id] int NOT NULL IDENTITY,
        [CierreProyectoRegistroId] int NOT NULL,
        [RecursoQAId] int NOT NULL,
        [NombreRecurso] nvarchar(200) NOT NULL,
        [Rol] nvarchar(100) NOT NULL,
        [Porcentaje] decimal(5,2) NOT NULL,
        [FechaInicio] datetime2 NOT NULL,
        [FechaFin] datetime2 NOT NULL,
        [Ejecutada] bit NOT NULL,
        CONSTRAINT [PK_CierreParticipantes] PRIMARY KEY ([Id]),
        CONSTRAINT [FK_CierreParticipantes_Cierres_CierreProyectoRegistroId] FOREIGN KEY ([CierreProyectoRegistroId]) REFERENCES [Cierres] ([Id]) ON DELETE CASCADE,
        CONSTRAINT [FK_CierreParticipantes_RecursosQA_RecursoQAId] FOREIGN KEY ([RecursoQAId]) REFERENCES [RecursosQA] ([Id]) ON DELETE NO ACTION
    );
END;
GO

IF NOT EXISTS (
    SELECT * FROM [__EFMigrationsHistory]
    WHERE [MigrationId] = N'20261009150116_Etapa9TrazabilidadCierre'
)
BEGIN
    CREATE INDEX [IX_ActividadAuditorias_ProyectoId_Fecha] ON [ActividadAuditorias] ([ProyectoId], [Fecha]);
END;
GO

IF NOT EXISTS (
    SELECT * FROM [__EFMigrationsHistory]
    WHERE [MigrationId] = N'20261009150116_Etapa9TrazabilidadCierre'
)
BEGIN
    CREATE INDEX [IX_CierreActividades_CierreProyectoRegistroId] ON [CierreActividades] ([CierreProyectoRegistroId]);
END;
GO

IF NOT EXISTS (
    SELECT * FROM [__EFMigrationsHistory]
    WHERE [MigrationId] = N'20261009150116_Etapa9TrazabilidadCierre'
)
BEGIN
    CREATE INDEX [IX_CierreParticipantes_CierreProyectoRegistroId] ON [CierreParticipantes] ([CierreProyectoRegistroId]);
END;
GO

IF NOT EXISTS (
    SELECT * FROM [__EFMigrationsHistory]
    WHERE [MigrationId] = N'20261009150116_Etapa9TrazabilidadCierre'
)
BEGIN
    CREATE INDEX [IX_CierreParticipantes_RecursoQAId] ON [CierreParticipantes] ([RecursoQAId]);
END;
GO

IF NOT EXISTS (
    SELECT * FROM [__EFMigrationsHistory]
    WHERE [MigrationId] = N'20261009150116_Etapa9TrazabilidadCierre'
)
BEGIN
    CREATE INDEX [IX_Cierres_ProyectoId] ON [Cierres] ([ProyectoId]);
END;
GO

IF NOT EXISTS (
    SELECT * FROM [__EFMigrationsHistory]
    WHERE [MigrationId] = N'20261009150116_Etapa9TrazabilidadCierre'
)
BEGIN

    INSERT INTO Cierres (ProyectoId, Fecha, Usuario, AvanceFinal, TotalActividades, ActividadesCompletadas, Reconstruido, Vigente)
    SELECT p.Id,
           ISNULL(DATEADD(DAY, 1, (SELECT MAX(a.FechaFin) FROM Asignaciones a WHERE a.ProyectoId = p.Id)), p.FechaFin),
           N'Reconstruido (migración)', 100,
           (SELECT COUNT(*) FROM Actividades x WHERE x.ProyectoId = p.Id),
           (SELECT COUNT(*) FROM Actividades x WHERE x.ProyectoId = p.Id AND x.Estado = 2),
           1, 1
    FROM Proyectos p
    WHERE p.Estado = 3 AND NOT EXISTS (SELECT 1 FROM Cierres c WHERE c.ProyectoId = p.Id);
END;
GO

IF NOT EXISTS (
    SELECT * FROM [__EFMigrationsHistory]
    WHERE [MigrationId] = N'20261009150116_Etapa9TrazabilidadCierre'
)
BEGIN

    INSERT INTO CierreParticipantes (CierreProyectoRegistroId, RecursoQAId, NombreRecurso, Rol, Porcentaje, FechaInicio, FechaFin, Ejecutada)
    SELECT c.Id, a.RecursoQAId, LTRIM(RTRIM(r.Nombres + N' ' + r.Apellidos)), a.Rol, a.Porcentaje, a.FechaInicio, a.FechaFin, 1
    FROM Cierres c
    JOIN Asignaciones a ON a.ProyectoId = c.ProyectoId
    JOIN RecursosQA r ON r.Id = a.RecursoQAId
    WHERE c.Reconstruido = 1;
END;
GO

IF NOT EXISTS (
    SELECT * FROM [__EFMigrationsHistory]
    WHERE [MigrationId] = N'20261009150116_Etapa9TrazabilidadCierre'
)
BEGIN

    INSERT INTO CierreActividades (CierreProyectoRegistroId, Nombre, Responsable, Estado, Avance, FechaInicio, FechaFin)
    SELECT c.Id, x.Nombre, NULL,
           CASE x.Estado WHEN 0 THEN N'Pendiente' WHEN 1 THEN N'En curso' WHEN 2 THEN N'Completada' ELSE N'Cancelada' END,
           x.Avance, x.FechaInicio, x.FechaFin
    FROM Cierres c
    JOIN Actividades x ON x.ProyectoId = c.ProyectoId
    WHERE c.Reconstruido = 1;
END;
GO

IF NOT EXISTS (
    SELECT * FROM [__EFMigrationsHistory]
    WHERE [MigrationId] = N'20261009150116_Etapa9TrazabilidadCierre'
)
BEGIN
    INSERT INTO [__EFMigrationsHistory] ([MigrationId], [ProductVersion])
    VALUES (N'20261009150116_Etapa9TrazabilidadCierre', N'8.0.11');
END;
GO

COMMIT;
GO

BEGIN TRANSACTION;
GO

IF NOT EXISTS (
    SELECT * FROM [__EFMigrationsHistory]
    WHERE [MigrationId] = N'20261009161934_Etapa10TipoAusencia'
)
BEGIN
    ALTER TABLE [Ausencias] ADD [Tipo] int NOT NULL DEFAULT 0;
END;
GO

IF NOT EXISTS (
    SELECT * FROM [__EFMigrationsHistory]
    WHERE [MigrationId] = N'20261009161934_Etapa10TipoAusencia'
)
BEGIN
    INSERT INTO [__EFMigrationsHistory] ([MigrationId], [ProductVersion])
    VALUES (N'20261009161934_Etapa10TipoAusencia', N'8.0.11');
END;
GO

COMMIT;
GO

BEGIN TRANSACTION;
GO

IF NOT EXISTS (
    SELECT * FROM [__EFMigrationsHistory]
    WHERE [MigrationId] = N'20261009185340_Etapa11ActividadesPredeterminadas'
)
BEGIN
    ALTER TABLE [Proyectos] ADD [AvanceLegado] decimal(5,2) NULL;
END;
GO

IF NOT EXISTS (
    SELECT * FROM [__EFMigrationsHistory]
    WHERE [MigrationId] = N'20261009185340_Etapa11ActividadesPredeterminadas'
)
BEGIN
    ALTER TABLE [Proyectos] ADD [AvancePendienteConfirmar] bit NOT NULL DEFAULT CAST(0 AS bit);
END;
GO

IF NOT EXISTS (
    SELECT * FROM [__EFMigrationsHistory]
    WHERE [MigrationId] = N'20261009185340_Etapa11ActividadesPredeterminadas'
)
BEGIN
    ALTER TABLE [Actividades] ADD [EsPredeterminada] bit NOT NULL DEFAULT CAST(0 AS bit);
END;
GO

IF NOT EXISTS (
    SELECT * FROM [__EFMigrationsHistory]
    WHERE [MigrationId] = N'20261009185340_Etapa11ActividadesPredeterminadas'
)
BEGIN
    ALTER TABLE [Actividades] ADD [Orden] int NOT NULL DEFAULT 0;
END;
GO

IF NOT EXISTS (
    SELECT * FROM [__EFMigrationsHistory]
    WHERE [MigrationId] = N'20261009185340_Etapa11ActividadesPredeterminadas'
)
BEGIN
    ALTER TABLE [Actividades] ADD [Peso] decimal(5,2) NOT NULL DEFAULT 10.0;
END;
GO

IF NOT EXISTS (
    SELECT * FROM [__EFMigrationsHistory]
    WHERE [MigrationId] = N'20261009185340_Etapa11ActividadesPredeterminadas'
)
BEGIN
    SELECT p.Id, p.Estado, p.Avance, CASE WHEN EXISTS
        (SELECT 1 FROM Actividades a WHERE a.ProyectoId=p.Id) THEN 1 ELSE 0 END ConActividades
    INTO #ProyectosEtapa11 FROM Proyectos p;

    UPDATE p SET AvancePendienteConfirmar=1, AvanceLegado=p.Avance
    FROM Proyectos p JOIN #ProyectosEtapa11 b ON b.Id=p.Id
    WHERE b.ConActividades=0 AND b.Estado=2;

    DECLARE @Def TABLE (Orden int, Nombre nvarchar(200), Peso decimal(5,2), Acumulado int);
    INSERT INTO @Def VALUES
    (1,N'Diseño y redacción de Casos de Prueba',30,30),
    (2,N'Ejecución de Casos de prueba',50,80),
    (3,N'Elaborar informe',20,100);

    -- Si hubiera nombres duplicados, se protege únicamente la primera coincidencia.
    ;WITH Coincidencias AS (
        SELECT a.Id,d.Nombre,d.Orden,d.Peso,ROW_NUMBER() OVER(PARTITION BY a.ProyectoId,d.Orden ORDER BY a.Id) rn
        FROM Actividades a JOIN @Def d ON LTRIM(RTRIM(a.Nombre))=d.Nombre
    )
    UPDATE a SET Nombre=c.Nombre,EsPredeterminada=1,Orden=c.Orden,Peso=c.Peso
    FROM Actividades a JOIN Coincidencias c ON c.Id=a.Id WHERE c.rn=1;
    UPDATE Actividades SET Orden=4 WHERE EsPredeterminada=0;
    UPDATE Actividades SET Estado=CASE WHEN Avance>=100 THEN 2 WHEN Avance>0 THEN 1 ELSE 0 END WHERE Estado<>3;

    ;WITH Rangos AS (
        SELECT p.Id,p.FechaInicio,p.FechaFin,b.Estado,b.ConActividades,
            CASE WHEN DATEDIFF(day,p.FechaInicio,p.FechaFin)<0 THEN 1
                 ELSE DATEDIFF(day,p.FechaInicio,p.FechaFin)+1 END Dias
        FROM Proyectos p JOIN #ProyectosEtapa11 b ON b.Id=p.Id
    ), Cortes AS (
        SELECT *,CASE WHEN Dias>=3 THEN
            CASE WHEN ROUND(Dias*0.3,0)<1 THEN 1 WHEN ROUND(Dias*0.3,0)>Dias-2 THEN Dias-2
                 ELSE CONVERT(int,ROUND(Dias*0.3,0)) END
            ELSE CONVERT(int,CEILING(Dias*0.3)) END C1 FROM Rangos
    ), Limites AS (
        SELECT *,CASE WHEN Dias>=3 THEN
            CASE WHEN ROUND(Dias*0.8,0)<C1+1 THEN C1+1 WHEN ROUND(Dias*0.8,0)>Dias-1 THEN Dias-1
                 ELSE CONVERT(int,ROUND(Dias*0.8,0)) END
            ELSE CONVERT(int,CEILING(Dias*0.8)) END C2 FROM Cortes
    )
    INSERT INTO Actividades (ProyectoId,Nombre,FechaInicio,FechaFin,Avance,Estado,RecursoQAId,EsPredeterminada,Orden,Peso)
    SELECT p.Id,d.Nombre,
        DATEADD(day,CASE d.Orden WHEN 1 THEN 0 WHEN 2 THEN CASE WHEN C1>=Dias THEN Dias-1 ELSE C1 END
            ELSE CASE WHEN C2>=Dias THEN Dias-1 ELSE C2 END END,CONVERT(date,p.FechaInicio)),
        DATEADD(day,CASE d.Orden WHEN 1 THEN C1-1 WHEN 2 THEN C2-1 ELSE Dias-1 END,CONVERT(date,p.FechaInicio)),
        CASE WHEN ConActividades=0 AND p.Estado=3 THEN 100 ELSE 0 END,
        CASE WHEN ConActividades=0 AND p.Estado=3 THEN 2 WHEN ConActividades=0 AND p.Estado=4 THEN 3 ELSE 0 END,
        NULL,1,d.Orden,d.Peso
    FROM Limites p CROSS JOIN @Def d
    WHERE NOT EXISTS (SELECT 1 FROM Actividades a WHERE a.ProyectoId=p.Id AND a.EsPredeterminada=1 AND a.Orden=d.Orden);

    -- El estado y sus efectos (foto de cierre / reapertura) se reconcilian con CierreProyecto
    -- al iniciar la aplicación o con el ejecutor de verificación, antes de reparar coherencia.
    UPDATE p SET Avance=CASE WHEN x.Promedio>=100 THEN 100
        WHEN ROUND(x.Promedio,2)>=100 THEN 99.99 ELSE ROUND(x.Promedio,2) END
    FROM Proyectos p JOIN #ProyectosEtapa11 b ON p.Id=b.Id
    CROSS APPLY (SELECT COALESCE(SUM(a.Avance*a.Peso)/NULLIF(SUM(a.Peso),0),0) Promedio
        FROM Actividades a WHERE a.ProyectoId=p.Id AND a.Estado<>3) x
    WHERE b.ConActividades=1 AND p.Estado<>4;
    DROP TABLE #ProyectosEtapa11;
END;
GO

IF NOT EXISTS (
    SELECT * FROM [__EFMigrationsHistory]
    WHERE [MigrationId] = N'20261009185340_Etapa11ActividadesPredeterminadas'
)
BEGIN
    INSERT INTO [__EFMigrationsHistory] ([MigrationId], [ProductVersion])
    VALUES (N'20261009185340_Etapa11ActividadesPredeterminadas', N'8.0.11');
END;
GO

COMMIT;
GO

BEGIN TRANSACTION;
GO

IF NOT EXISTS (
    SELECT * FROM [__EFMigrationsHistory]
    WHERE [MigrationId] = N'20261012165839_Etapa12Acuerdos'
)
BEGIN
    CREATE TABLE [Acuerdos] (
        [Id] int NOT NULL IDENTITY,
        [ProyectoId] int NOT NULL,
        [Numero] int NOT NULL,
        [Titulo] nvarchar(200) NOT NULL,
        [RecursoQAId] int NULL,
        [ResponsableNombre] nvarchar(150) NULL,
        [FechaAcuerdo] datetime2 NOT NULL,
        [FechaLimite] datetime2 NOT NULL,
        [Estado] int NOT NULL,
        CONSTRAINT [PK_Acuerdos] PRIMARY KEY ([Id]),
        CONSTRAINT [FK_Acuerdos_Proyectos_ProyectoId] FOREIGN KEY ([ProyectoId]) REFERENCES [Proyectos] ([Id]) ON DELETE CASCADE,
        CONSTRAINT [FK_Acuerdos_RecursosQA_RecursoQAId] FOREIGN KEY ([RecursoQAId]) REFERENCES [RecursosQA] ([Id]) ON DELETE NO ACTION
    );
END;
GO

IF NOT EXISTS (
    SELECT * FROM [__EFMigrationsHistory]
    WHERE [MigrationId] = N'20261012165839_Etapa12Acuerdos'
)
BEGIN
    CREATE TABLE [AcuerdoActividades] (
        [AcuerdoId] int NOT NULL,
        [ActividadId] int NOT NULL,
        CONSTRAINT [PK_AcuerdoActividades] PRIMARY KEY ([AcuerdoId], [ActividadId]),
        CONSTRAINT [FK_AcuerdoActividades_Actividades_ActividadId] FOREIGN KEY ([ActividadId]) REFERENCES [Actividades] ([Id]) ON DELETE NO ACTION,
        CONSTRAINT [FK_AcuerdoActividades_Acuerdos_AcuerdoId] FOREIGN KEY ([AcuerdoId]) REFERENCES [Acuerdos] ([Id]) ON DELETE CASCADE
    );
END;
GO

IF NOT EXISTS (
    SELECT * FROM [__EFMigrationsHistory]
    WHERE [MigrationId] = N'20261012165839_Etapa12Acuerdos'
)
BEGIN
    CREATE INDEX [IX_AcuerdoActividades_ActividadId] ON [AcuerdoActividades] ([ActividadId]);
END;
GO

IF NOT EXISTS (
    SELECT * FROM [__EFMigrationsHistory]
    WHERE [MigrationId] = N'20261012165839_Etapa12Acuerdos'
)
BEGIN
    CREATE UNIQUE INDEX [IX_Acuerdos_ProyectoId_Numero] ON [Acuerdos] ([ProyectoId], [Numero]);
END;
GO

IF NOT EXISTS (
    SELECT * FROM [__EFMigrationsHistory]
    WHERE [MigrationId] = N'20261012165839_Etapa12Acuerdos'
)
BEGIN
    CREATE INDEX [IX_Acuerdos_RecursoQAId] ON [Acuerdos] ([RecursoQAId]);
END;
GO

IF NOT EXISTS (
    SELECT * FROM [__EFMigrationsHistory]
    WHERE [MigrationId] = N'20261012165839_Etapa12Acuerdos'
)
BEGIN
    INSERT INTO [__EFMigrationsHistory] ([MigrationId], [ProductVersion])
    VALUES (N'20261012165839_Etapa12Acuerdos', N'8.0.11');
END;
GO

COMMIT;
GO

BEGIN TRANSACTION;
GO

IF NOT EXISTS (
    SELECT * FROM [__EFMigrationsHistory]
    WHERE [MigrationId] = N'20261012190345_Etapa13FuenteAvanceActividad'
)
BEGIN
    ALTER TABLE [Actividades] ADD [FuenteAvance] int NOT NULL DEFAULT 0;
END;
GO

IF NOT EXISTS (
    SELECT * FROM [__EFMigrationsHistory]
    WHERE [MigrationId] = N'20261012190345_Etapa13FuenteAvanceActividad'
)
BEGIN
    INSERT INTO [__EFMigrationsHistory] ([MigrationId], [ProductVersion])
    VALUES (N'20261012190345_Etapa13FuenteAvanceActividad', N'8.0.11');
END;
GO

COMMIT;
GO

BEGIN TRANSACTION;
GO

IF NOT EXISTS (
    SELECT * FROM [__EFMigrationsHistory]
    WHERE [MigrationId] = N'20261012195213_Etapa14SquashProyectoVinculado'
)
BEGIN
    DECLARE @var1 sysname;
    SELECT @var1 = [d].[name]
    FROM [sys].[default_constraints] [d]
    INNER JOIN [sys].[columns] [c] ON [d].[parent_column_id] = [c].[column_id] AND [d].[parent_object_id] = [c].[object_id]
    WHERE ([d].[parent_object_id] = OBJECT_ID(N'[Actividades]') AND [c].[name] = N'FuenteAvance');
    IF @var1 IS NOT NULL EXEC(N'ALTER TABLE [Actividades] DROP CONSTRAINT [' + @var1 + '];');
    ALTER TABLE [Actividades] DROP COLUMN [FuenteAvance];
END;
GO

IF NOT EXISTS (
    SELECT * FROM [__EFMigrationsHistory]
    WHERE [MigrationId] = N'20261012195213_Etapa14SquashProyectoVinculado'
)
BEGIN
    ALTER TABLE [RecursosQA] ADD [SquashLogin] nvarchar(100) NULL;
END;
GO

IF NOT EXISTS (
    SELECT * FROM [__EFMigrationsHistory]
    WHERE [MigrationId] = N'20261012195213_Etapa14SquashProyectoVinculado'
)
BEGIN
    ALTER TABLE [Proyectos] ADD [SquashProjectId] int NULL;
END;
GO

IF NOT EXISTS (
    SELECT * FROM [__EFMigrationsHistory]
    WHERE [MigrationId] = N'20261012195213_Etapa14SquashProyectoVinculado'
)
BEGIN
    ALTER TABLE [Proyectos] ADD [SquashSincronizadoEn] datetime2 NULL;
END;
GO

IF NOT EXISTS (
    SELECT * FROM [__EFMigrationsHistory]
    WHERE [MigrationId] = N'20261012195213_Etapa14SquashProyectoVinculado'
)
BEGIN
    CREATE TABLE [SquashCampanas] (
        [Id] int NOT NULL IDENTITY,
        [ProyectoId] int NOT NULL,
        [SquashId] int NOT NULL,
        [Nombre] nvarchar(200) NOT NULL,
        [Estado] nvarchar(30) NULL,
        [Vigente] bit NOT NULL,
        CONSTRAINT [PK_SquashCampanas] PRIMARY KEY ([Id]),
        CONSTRAINT [FK_SquashCampanas_Proyectos_ProyectoId] FOREIGN KEY ([ProyectoId]) REFERENCES [Proyectos] ([Id]) ON DELETE CASCADE
    );
END;
GO

IF NOT EXISTS (
    SELECT * FROM [__EFMigrationsHistory]
    WHERE [MigrationId] = N'20261012195213_Etapa14SquashProyectoVinculado'
)
BEGIN
    CREATE TABLE [SquashIteraciones] (
        [Id] int NOT NULL IDENTITY,
        [SquashCampanaId] int NOT NULL,
        [SquashId] int NOT NULL,
        [Nombre] nvarchar(200) NOT NULL,
        [Estado] nvarchar(30) NULL,
        [Porcentaje] decimal(5,2) NOT NULL,
        [Vigente] bit NOT NULL,
        CONSTRAINT [PK_SquashIteraciones] PRIMARY KEY ([Id]),
        CONSTRAINT [FK_SquashIteraciones_SquashCampanas_SquashCampanaId] FOREIGN KEY ([SquashCampanaId]) REFERENCES [SquashCampanas] ([Id]) ON DELETE CASCADE
    );
END;
GO

IF NOT EXISTS (
    SELECT * FROM [__EFMigrationsHistory]
    WHERE [MigrationId] = N'20261012195213_Etapa14SquashProyectoVinculado'
)
BEGIN
    CREATE TABLE [SquashCasos] (
        [Id] int NOT NULL IDENTITY,
        [SquashIteracionId] int NOT NULL,
        [SquashItemId] int NOT NULL,
        [Nombre] nvarchar(300) NOT NULL,
        [EstadoSquash] nvarchar(30) NOT NULL,
        [Porcentaje] decimal(5,2) NULL,
        [AsignadoLogin] nvarchar(100) NULL,
        [RecursoQAId] int NULL,
        [UltimaEjecucion] datetime2 NULL,
        [Vigente] bit NOT NULL,
        CONSTRAINT [PK_SquashCasos] PRIMARY KEY ([Id]),
        CONSTRAINT [FK_SquashCasos_RecursosQA_RecursoQAId] FOREIGN KEY ([RecursoQAId]) REFERENCES [RecursosQA] ([Id]) ON DELETE NO ACTION,
        CONSTRAINT [FK_SquashCasos_SquashIteraciones_SquashIteracionId] FOREIGN KEY ([SquashIteracionId]) REFERENCES [SquashIteraciones] ([Id]) ON DELETE CASCADE
    );
END;
GO

IF NOT EXISTS (
    SELECT * FROM [__EFMigrationsHistory]
    WHERE [MigrationId] = N'20261012195213_Etapa14SquashProyectoVinculado'
)
BEGIN
    UPDATE RecursosQA SET SquashLogin = Correo WHERE SquashLogin IS NULL
END;
GO

IF NOT EXISTS (
    SELECT * FROM [__EFMigrationsHistory]
    WHERE [MigrationId] = N'20261012195213_Etapa14SquashProyectoVinculado'
)
BEGIN
    EXEC(N'CREATE UNIQUE INDEX [IX_RecursosQA_SquashLogin] ON [RecursosQA] ([SquashLogin]) WHERE [SquashLogin] IS NOT NULL');
END;
GO

IF NOT EXISTS (
    SELECT * FROM [__EFMigrationsHistory]
    WHERE [MigrationId] = N'20261012195213_Etapa14SquashProyectoVinculado'
)
BEGIN
    EXEC(N'CREATE UNIQUE INDEX [IX_Proyectos_SquashProjectId] ON [Proyectos] ([SquashProjectId]) WHERE [SquashProjectId] IS NOT NULL');
END;
GO

IF NOT EXISTS (
    SELECT * FROM [__EFMigrationsHistory]
    WHERE [MigrationId] = N'20261012195213_Etapa14SquashProyectoVinculado'
)
BEGIN
    CREATE UNIQUE INDEX [IX_SquashCampanas_ProyectoId_SquashId] ON [SquashCampanas] ([ProyectoId], [SquashId]);
END;
GO

IF NOT EXISTS (
    SELECT * FROM [__EFMigrationsHistory]
    WHERE [MigrationId] = N'20261012195213_Etapa14SquashProyectoVinculado'
)
BEGIN
    CREATE INDEX [IX_SquashCasos_RecursoQAId] ON [SquashCasos] ([RecursoQAId]);
END;
GO

IF NOT EXISTS (
    SELECT * FROM [__EFMigrationsHistory]
    WHERE [MigrationId] = N'20261012195213_Etapa14SquashProyectoVinculado'
)
BEGIN
    CREATE UNIQUE INDEX [IX_SquashCasos_SquashIteracionId_SquashItemId] ON [SquashCasos] ([SquashIteracionId], [SquashItemId]);
END;
GO

IF NOT EXISTS (
    SELECT * FROM [__EFMigrationsHistory]
    WHERE [MigrationId] = N'20261012195213_Etapa14SquashProyectoVinculado'
)
BEGIN
    CREATE UNIQUE INDEX [IX_SquashIteraciones_SquashCampanaId_SquashId] ON [SquashIteraciones] ([SquashCampanaId], [SquashId]);
END;
GO

IF NOT EXISTS (
    SELECT * FROM [__EFMigrationsHistory]
    WHERE [MigrationId] = N'20261012195213_Etapa14SquashProyectoVinculado'
)
BEGIN
    INSERT INTO [__EFMigrationsHistory] ([MigrationId], [ProductVersion])
    VALUES (N'20261012195213_Etapa14SquashProyectoVinculado', N'8.0.11');
END;
GO

COMMIT;
GO

BEGIN TRANSACTION;
GO

IF NOT EXISTS (
    SELECT * FROM [__EFMigrationsHistory]
    WHERE [MigrationId] = N'20261012231457_Etapa15BpoProyecto'
)
BEGIN
    ALTER TABLE [Proyectos] ADD [BPO] nvarchar(150) NULL;
END;
GO

IF NOT EXISTS (
    SELECT * FROM [__EFMigrationsHistory]
    WHERE [MigrationId] = N'20261012231457_Etapa15BpoProyecto'
)
BEGIN
    INSERT INTO [__EFMigrationsHistory] ([MigrationId], [ProductVersion])
    VALUES (N'20261012231457_Etapa15BpoProyecto', N'8.0.11');
END;
GO

COMMIT;
GO

BEGIN TRANSACTION;
GO

IF NOT EXISTS (
    SELECT * FROM [__EFMigrationsHistory]
    WHERE [MigrationId] = N'20261012231814_Etapa16SquashTestCaseId'
)
BEGIN
    ALTER TABLE [SquashCasos] ADD [SquashTestCaseId] int NULL;
END;
GO

IF NOT EXISTS (
    SELECT * FROM [__EFMigrationsHistory]
    WHERE [MigrationId] = N'20261012231814_Etapa16SquashTestCaseId'
)
BEGIN
    INSERT INTO [__EFMigrationsHistory] ([MigrationId], [ProductVersion])
    VALUES (N'20261012231814_Etapa16SquashTestCaseId', N'8.0.11');
END;
GO

COMMIT;
GO

BEGIN TRANSACTION;
GO

IF NOT EXISTS (
    SELECT * FROM [__EFMigrationsHistory]
    WHERE [MigrationId] = N'20261016203429_Etapa17VinculoBugzilla'
)
BEGIN
    CREATE TABLE [VinculosBugzilla] (
        [Id] int NOT NULL IDENTITY,
        [ProyectoId] int NOT NULL,
        [SquashAmbiente] nvarchar(300) NOT NULL,
        [SquashProjectId] int NOT NULL,
        [BugzillaAmbiente] nvarchar(300) NOT NULL,
        [ProductoId] int NULL,
        [ProductoNombre] nvarchar(200) NULL,
        [ComponenteId] int NULL,
        [ComponenteNombre] nvarchar(100) NULL,
        [VersionId] int NULL,
        [VersionNombre] nvarchar(100) NULL,
        [Estado] int NOT NULL,
        [UltimaSincronizacion] datetime2 NULL,
        [UltimoIntento] datetime2 NULL,
        [UltimoError] nvarchar(500) NULL,
        [AsociacionExistenteConfirmada] bit NOT NULL,
        [BloqueadoHasta] datetime2 NULL,
        [CreadoEn] datetime2 NOT NULL,
        CONSTRAINT [PK_VinculosBugzilla] PRIMARY KEY ([Id]),
        CONSTRAINT [FK_VinculosBugzilla_Proyectos_ProyectoId] FOREIGN KEY ([ProyectoId]) REFERENCES [Proyectos] ([Id]) ON DELETE CASCADE
    );
END;
GO

IF NOT EXISTS (
    SELECT * FROM [__EFMigrationsHistory]
    WHERE [MigrationId] = N'20261016203429_Etapa17VinculoBugzilla'
)
BEGIN
    EXEC(N'CREATE UNIQUE INDEX [IX_VinculosBugzilla_BugzillaAmbiente_ProductoId] ON [VinculosBugzilla] ([BugzillaAmbiente], [ProductoId]) WHERE [ProductoId] IS NOT NULL');
END;
GO

IF NOT EXISTS (
    SELECT * FROM [__EFMigrationsHistory]
    WHERE [MigrationId] = N'20261016203429_Etapa17VinculoBugzilla'
)
BEGIN
    CREATE UNIQUE INDEX [IX_VinculosBugzilla_ProyectoId] ON [VinculosBugzilla] ([ProyectoId]);
END;
GO

IF NOT EXISTS (
    SELECT * FROM [__EFMigrationsHistory]
    WHERE [MigrationId] = N'20261016203429_Etapa17VinculoBugzilla'
)
BEGIN
    CREATE UNIQUE INDEX [IX_VinculosBugzilla_SquashAmbiente_SquashProjectId] ON [VinculosBugzilla] ([SquashAmbiente], [SquashProjectId]);
END;
GO

IF NOT EXISTS (
    SELECT * FROM [__EFMigrationsHistory]
    WHERE [MigrationId] = N'20261016203429_Etapa17VinculoBugzilla'
)
BEGIN
    INSERT INTO [__EFMigrationsHistory] ([MigrationId], [ProductVersion])
    VALUES (N'20261016203429_Etapa17VinculoBugzilla', N'8.0.11');
END;
GO

COMMIT;
GO

BEGIN TRANSACTION;
GO

IF NOT EXISTS (
    SELECT * FROM [__EFMigrationsHistory]
    WHERE [MigrationId] = N'20261017003925_Etapa18SquashModulos'
)
BEGIN
    CREATE TABLE [SquashModulos] (
        [Id] int NOT NULL IDENTITY,
        [SquashIteracionId] int NOT NULL,
        [SquashId] int NOT NULL,
        [Nombre] nvarchar(200) NOT NULL,
        [Vigente] bit NOT NULL,
        CONSTRAINT [PK_SquashModulos] PRIMARY KEY ([Id]),
        CONSTRAINT [FK_SquashModulos_SquashIteraciones_SquashIteracionId] FOREIGN KEY ([SquashIteracionId]) REFERENCES [SquashIteraciones] ([Id]) ON DELETE CASCADE
    );
END;
GO

IF NOT EXISTS (
    SELECT * FROM [__EFMigrationsHistory]
    WHERE [MigrationId] = N'20261017003925_Etapa18SquashModulos'
)
BEGIN
    CREATE TABLE [SquashModuloCaso] (
        [CasoId] int NOT NULL,
        [ModuloId] int NOT NULL,
        CONSTRAINT [PK_SquashModuloCaso] PRIMARY KEY ([CasoId], [ModuloId]),
        CONSTRAINT [FK_SquashModuloCaso_SquashCasos_CasoId] FOREIGN KEY ([CasoId]) REFERENCES [SquashCasos] ([Id]),
        CONSTRAINT [FK_SquashModuloCaso_SquashModulos_ModuloId] FOREIGN KEY ([ModuloId]) REFERENCES [SquashModulos] ([Id]) ON DELETE CASCADE
    );
END;
GO

IF NOT EXISTS (
    SELECT * FROM [__EFMigrationsHistory]
    WHERE [MigrationId] = N'20261017003925_Etapa18SquashModulos'
)
BEGIN
    CREATE INDEX [IX_SquashModuloCaso_ModuloId] ON [SquashModuloCaso] ([ModuloId]);
END;
GO

IF NOT EXISTS (
    SELECT * FROM [__EFMigrationsHistory]
    WHERE [MigrationId] = N'20261017003925_Etapa18SquashModulos'
)
BEGIN
    CREATE UNIQUE INDEX [IX_SquashModulos_SquashIteracionId_SquashId] ON [SquashModulos] ([SquashIteracionId], [SquashId]);
END;
GO

IF NOT EXISTS (
    SELECT * FROM [__EFMigrationsHistory]
    WHERE [MigrationId] = N'20261017003925_Etapa18SquashModulos'
)
BEGIN
    INSERT INTO [__EFMigrationsHistory] ([MigrationId], [ProductVersion])
    VALUES (N'20261017003925_Etapa18SquashModulos', N'8.0.11');
END;
GO

COMMIT;
GO


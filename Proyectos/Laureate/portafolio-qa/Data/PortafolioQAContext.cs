using Microsoft.AspNetCore.Identity.EntityFrameworkCore;
using Microsoft.EntityFrameworkCore;
using PortafolioQA.Models;

namespace PortafolioQA.Data;

public class PortafolioQAContext : IdentityDbContext<UsuarioApp>
{
    public PortafolioQAContext(DbContextOptions<PortafolioQAContext> options) : base(options)
    {
    }

    public DbSet<Institucion> Instituciones => Set<Institucion>();
    public DbSet<Sistema> Sistemas => Set<Sistema>();
    public DbSet<SistemaInstitucion> SistemaInstituciones => Set<SistemaInstitucion>();
    public DbSet<RecursoQA> RecursosQA => Set<RecursoQA>();
    public DbSet<Proyecto> Proyectos => Set<Proyecto>();
    public DbSet<ProyectoSistema> ProyectoSistemas => Set<ProyectoSistema>();
    public DbSet<Asignacion> Asignaciones => Set<Asignacion>();
    public DbSet<AsociacionExterna> AsociacionesExternas => Set<AsociacionExterna>();
    public DbSet<Feriado> Feriados => Set<Feriado>();
    public DbSet<Ausencia> Ausencias => Set<Ausencia>();
    public DbSet<Hito> Hitos => Set<Hito>();
    public DbSet<Actividad> Actividades => Set<Actividad>();
    public DbSet<CierreProyectoRegistro> Cierres => Set<CierreProyectoRegistro>();
    public DbSet<CierreParticipante> CierreParticipantes => Set<CierreParticipante>();
    public DbSet<CierreActividad> CierreActividades => Set<CierreActividad>();
    public DbSet<ActividadAuditoria> ActividadAuditorias => Set<ActividadAuditoria>();
    public DbSet<Acuerdo> Acuerdos => Set<Acuerdo>();
    public DbSet<AcuerdoActividad> AcuerdoActividades => Set<AcuerdoActividad>();
    public DbSet<VinculoBugzilla> VinculosBugzilla => Set<VinculoBugzilla>();
    public DbSet<SquashCampana> SquashCampanas => Set<SquashCampana>();
    public DbSet<SquashIteracion> SquashIteraciones => Set<SquashIteracion>();
    public DbSet<SquashCaso> SquashCasos => Set<SquashCaso>();
    public DbSet<SquashModulo> SquashModulos => Set<SquashModulo>();

    private void ValidarActividades()
    {
        ChangeTracker.DetectChanges();
        foreach (var entry in ChangeTracker.Entries<Actividad>())
        {
            if (entry.State == EntityState.Modified && entry.OriginalValues.GetValue<bool>(nameof(Actividad.EsPredeterminada)))
            {
                if (entry.Property(a => a.Nombre).IsModified || entry.Property(a => a.EsPredeterminada).IsModified || entry.Property(a => a.Orden).IsModified)
                    throw new InvalidOperationException("Las actividades predeterminadas no se pueden renombrar ni desproteger.");
            }
            if (entry.State == EntityState.Deleted && entry.Entity.EsPredeterminada
                && !ChangeTracker.Entries<Proyecto>().Any(p => p.Entity.Id == entry.Entity.ProyectoId && p.State == EntityState.Deleted))
                throw new InvalidOperationException("Las actividades predeterminadas no se pueden eliminar.");
            if (entry.State is EntityState.Added or EntityState.Modified && entry.Entity.Peso is < 0.01m or > 999.99m)
                throw new InvalidOperationException("El peso debe estar entre 0.01 y 999.99.");
        }
    }

    public override int SaveChanges(bool acceptAllChangesOnSuccess)
    {
        ValidarActividades();
        return base.SaveChanges(acceptAllChangesOnSuccess);
    }

    public override Task<int> SaveChangesAsync(bool acceptAllChangesOnSuccess, CancellationToken cancellationToken = default)
    {
        ValidarActividades();
        return base.SaveChangesAsync(acceptAllChangesOnSuccess, cancellationToken);
    }

    protected override void OnModelCreating(ModelBuilder modelBuilder)
    {
        base.OnModelCreating(modelBuilder);

        modelBuilder.Entity<SistemaInstitucion>()
            .HasKey(si => new { si.SistemaId, si.InstitucionId });
        modelBuilder.Entity<SistemaInstitucion>()
            .HasOne(si => si.Sistema)
            .WithMany(s => s.Instituciones)
            .HasForeignKey(si => si.SistemaId);
        modelBuilder.Entity<SistemaInstitucion>()
            .HasOne(si => si.Institucion)
            .WithMany(i => i.Sistemas)
            .HasForeignKey(si => si.InstitucionId);

        modelBuilder.Entity<ProyectoSistema>()
            .HasKey(ps => new { ps.ProyectoId, ps.SistemaId });
        modelBuilder.Entity<ProyectoSistema>()
            .HasOne(ps => ps.Proyecto)
            .WithMany(p => p.Sistemas)
            .HasForeignKey(ps => ps.ProyectoId);
        modelBuilder.Entity<ProyectoSistema>()
            .HasOne(ps => ps.Sistema)
            .WithMany(s => s.Proyectos)
            .HasForeignKey(ps => ps.SistemaId);

        modelBuilder.Entity<Institucion>()
            .HasIndex(i => i.Nombre)
            .IsUnique();

        modelBuilder.Entity<Sistema>()
            .HasIndex(s => s.Nombre)
            .IsUnique();

        modelBuilder.Entity<RecursoQA>()
            .HasIndex(r => r.Correo)
            .IsUnique();

        modelBuilder.Entity<RecursoQA>()
            .HasOne(r => r.Usuario)
            .WithMany()
            .HasForeignKey(r => r.UsuarioId)
            .OnDelete(DeleteBehavior.Restrict);

        modelBuilder.Entity<RecursoQA>()
            .HasIndex(r => r.UsuarioId)
            .IsUnique()
            .HasFilter("[UsuarioId] IS NOT NULL");

        modelBuilder.Entity<Proyecto>()
            .HasOne(p => p.Institucion)
            .WithMany(i => i.Proyectos)
            .HasForeignKey(p => p.InstitucionId)
            .OnDelete(DeleteBehavior.Restrict);

        modelBuilder.Entity<Asignacion>()
            .HasOne(a => a.Proyecto)
            .WithMany(p => p.Asignaciones)
            .HasForeignKey(a => a.ProyectoId)
            .OnDelete(DeleteBehavior.Cascade);

        modelBuilder.Entity<Asignacion>()
            .HasOne(a => a.RecursoQA)
            .WithMany(r => r.Asignaciones)
            .HasForeignKey(a => a.RecursoQAId)
            .OnDelete(DeleteBehavior.Restrict);

        modelBuilder.Entity<Feriado>()
            .HasIndex(f => f.Fecha)
            .IsUnique();

        modelBuilder.Entity<Ausencia>()
            .HasOne(a => a.RecursoQA)
            .WithMany()
            .HasForeignKey(a => a.RecursoQAId)
            .OnDelete(DeleteBehavior.Restrict);

        modelBuilder.Entity<Hito>()
            .HasOne(h => h.Proyecto)
            .WithMany(p => p.Hitos)
            .HasForeignKey(h => h.ProyectoId)
            .OnDelete(DeleteBehavior.Cascade);

        modelBuilder.Entity<Actividad>().Property(a => a.Peso).HasDefaultValue(10m);

        modelBuilder.Entity<Actividad>()
            .HasOne(a => a.Proyecto)
            .WithMany(p => p.Actividades)
            .HasForeignKey(a => a.ProyectoId)
            .OnDelete(DeleteBehavior.Cascade);

        modelBuilder.Entity<Actividad>()
            .HasOne(a => a.RecursoQA)
            .WithMany()
            .HasForeignKey(a => a.RecursoQAId)
            .OnDelete(DeleteBehavior.Restrict);

        modelBuilder.Entity<AsociacionExterna>()
            .HasOne(a => a.Proyecto)
            .WithMany(p => p.Integraciones)
            .HasForeignKey(a => a.ProyectoId)
            .OnDelete(DeleteBehavior.Cascade);

        modelBuilder.Entity<CierreProyectoRegistro>()
            .HasOne(c => c.Proyecto).WithMany(p => p.Cierres)
            .HasForeignKey(c => c.ProyectoId).OnDelete(DeleteBehavior.Cascade);
        modelBuilder.Entity<CierreParticipante>()
            .HasOne(c => c.Cierre).WithMany(r => r.Participantes)
            .HasForeignKey(c => c.CierreProyectoRegistroId).OnDelete(DeleteBehavior.Cascade);
        modelBuilder.Entity<CierreParticipante>()
            .HasOne(c => c.RecursoQA).WithMany()
            .HasForeignKey(c => c.RecursoQAId).OnDelete(DeleteBehavior.Restrict);
        modelBuilder.Entity<CierreActividad>()
            .HasOne(c => c.Cierre).WithMany(r => r.Actividades)
            .HasForeignKey(c => c.CierreProyectoRegistroId).OnDelete(DeleteBehavior.Cascade);
        modelBuilder.Entity<ActividadAuditoria>()
            .HasOne(a => a.Proyecto).WithMany()
            .HasForeignKey(a => a.ProyectoId).OnDelete(DeleteBehavior.Cascade);
        modelBuilder.Entity<ActividadAuditoria>().HasIndex(a => new { a.ProyectoId, a.Fecha });

        modelBuilder.Entity<Acuerdo>()
            .HasOne(a => a.Proyecto).WithMany(p => p.Acuerdos)
            .HasForeignKey(a => a.ProyectoId).OnDelete(DeleteBehavior.Cascade);
        modelBuilder.Entity<Acuerdo>()
            .HasOne(a => a.RecursoQA).WithMany()
            .HasForeignKey(a => a.RecursoQAId).OnDelete(DeleteBehavior.Restrict);
        modelBuilder.Entity<Acuerdo>().HasIndex(a => new { a.ProyectoId, a.Numero }).IsUnique();

        // SQL Server no admite dos rutas de cascada (Proyecto→Actividad y Proyecto→Acuerdo): el vínculo cae con el
        // acuerdo; al eliminar una actividad se quitan sus vínculos explícitamente (Edit.cshtml.cs).
        modelBuilder.Entity<Proyecto>().HasIndex(p => p.SquashProjectId).IsUnique().HasFilter("[SquashProjectId] IS NOT NULL");
        modelBuilder.Entity<RecursoQA>().HasIndex(r => r.SquashLogin).IsUnique().HasFilter("[SquashLogin] IS NOT NULL");

        modelBuilder.Entity<VinculoBugzilla>(v =>
        {
            v.HasOne(x => x.Proyecto).WithOne(p => p.VinculoBugzilla).HasForeignKey<VinculoBugzilla>(x => x.ProyectoId).OnDelete(DeleteBehavior.Cascade);
            v.HasIndex(x => x.ProyectoId).IsUnique();
            // Un proyecto de Squash y un producto de Bugzilla solo pueden pertenecer a una iniciativa, siempre dentro de su ambiente.
            v.HasIndex(x => new { x.SquashAmbiente, x.SquashProjectId }).IsUnique();
            v.HasIndex(x => new { x.BugzillaAmbiente, x.ProductoId }).IsUnique().HasFilter("[ProductoId] IS NOT NULL");
        });

        modelBuilder.Entity<SquashCampana>()
            .HasOne(c => c.Proyecto).WithMany(p => p.SquashCampanas)
            .HasForeignKey(c => c.ProyectoId).OnDelete(DeleteBehavior.Cascade);
        modelBuilder.Entity<SquashCampana>().HasIndex(c => new { c.ProyectoId, c.SquashId }).IsUnique();
        modelBuilder.Entity<SquashIteracion>()
            .HasOne(i => i.SquashCampana).WithMany(c => c.Iteraciones)
            .HasForeignKey(i => i.SquashCampanaId).OnDelete(DeleteBehavior.Cascade);
        modelBuilder.Entity<SquashIteracion>().HasIndex(i => new { i.SquashCampanaId, i.SquashId }).IsUnique();
        modelBuilder.Entity<SquashIteracion>().Property(i => i.Porcentaje).HasPrecision(5, 2);
        modelBuilder.Entity<SquashCaso>()
            .HasOne(k => k.SquashIteracion).WithMany(i => i.Casos)
            .HasForeignKey(k => k.SquashIteracionId).OnDelete(DeleteBehavior.Cascade);
        modelBuilder.Entity<SquashCaso>()
            .HasOne(k => k.RecursoQA).WithMany()
            .HasForeignKey(k => k.RecursoQAId).OnDelete(DeleteBehavior.Restrict);
        modelBuilder.Entity<SquashCaso>().HasIndex(k => new { k.SquashIteracionId, k.SquashItemId }).IsUnique();
        modelBuilder.Entity<SquashCaso>().Property(k => k.Porcentaje).HasPrecision(5, 2);
        modelBuilder.Entity<SquashModulo>()
            .HasOne(m => m.SquashIteracion).WithMany(i => i.Modulos)
            .HasForeignKey(m => m.SquashIteracionId).OnDelete(DeleteBehavior.Cascade);
        modelBuilder.Entity<SquashModulo>().HasIndex(m => new { m.SquashIteracionId, m.SquashId }).IsUnique();
        // Relación módulo ↔ ítem. La tabla de unión borra en cascada por el módulo; por el caso no (SQL Server no admite dos rutas en cascada).
        modelBuilder.Entity<SquashModulo>()
            .HasMany(m => m.Casos).WithMany(k => k.Modulos)
            .UsingEntity<Dictionary<string, object>>("SquashModuloCaso",
                j => j.HasOne<SquashCaso>().WithMany().HasForeignKey("CasoId").OnDelete(DeleteBehavior.NoAction),
                j => j.HasOne<SquashModulo>().WithMany().HasForeignKey("ModuloId").OnDelete(DeleteBehavior.Cascade));

        modelBuilder.Entity<AcuerdoActividad>().HasKey(x => new { x.AcuerdoId, x.ActividadId });
        modelBuilder.Entity<AcuerdoActividad>()
            .HasOne(x => x.Acuerdo).WithMany(a => a.Actividades)
            .HasForeignKey(x => x.AcuerdoId).OnDelete(DeleteBehavior.Cascade);
        modelBuilder.Entity<AcuerdoActividad>()
            .HasOne(x => x.Actividad).WithMany()
            .HasForeignKey(x => x.ActividadId).OnDelete(DeleteBehavior.Restrict);

        foreach (var decimalProperty in modelBuilder.Model.GetEntityTypes()
                     .SelectMany(t => t.GetProperties())
                     .Where(p => p.ClrType == typeof(decimal) || p.ClrType == typeof(decimal?)))
        {
            decimalProperty.SetColumnType("decimal(5,2)");
        }
    }
}

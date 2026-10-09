using System.ComponentModel.DataAnnotations;

namespace PortafolioQA.Models;

public class Sistema
{
    public int Id { get; set; }

    [Required(ErrorMessage = "El nombre es obligatorio.")]
    [StringLength(150)]
    public string Nombre { get; set; } = string.Empty;

    public bool Activo { get; set; } = true;

    public List<SistemaInstitucion> Instituciones { get; set; } = new();
    public List<ProyectoSistema> Proyectos { get; set; } = new();
}

public class SistemaInstitucion
{
    public int SistemaId { get; set; }
    public Sistema Sistema { get; set; } = null!;

    public int InstitucionId { get; set; }
    public Institucion Institucion { get; set; } = null!;
}

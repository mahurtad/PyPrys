using System.ComponentModel.DataAnnotations;

namespace PortafolioQA.Models;

public class Institucion
{
    public int Id { get; set; }

    [Required(ErrorMessage = "El nombre es obligatorio.")]
    [StringLength(150)]
    public string Nombre { get; set; } = string.Empty;

    public bool Activo { get; set; } = true;

    public List<Proyecto> Proyectos { get; set; } = new();
    public List<SistemaInstitucion> Sistemas { get; set; } = new();
}

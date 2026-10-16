using PortafolioQA.Models;

namespace PortafolioQA.Services;

/// <summary>
/// Un mismo caso de prueba de Squash puede aparecer en varios ítems (varias iteraciones, campañas o repetido en una iteración).
/// Para certificar se cuenta cada caso una sola vez, con su resultado más reciente.
/// </summary>
public static class SquashCasosUnicos
{
    /// <summary>Un caso por SquashTestCaseId (o por ítem si no se conoce); gana la última ejecución y, en empate, el ítem más nuevo.</summary>
    public static List<SquashCaso> Depurar(IEnumerable<SquashCaso> casos) =>
        casos.GroupBy(k => k.SquashTestCaseId is int tc ? $"tc:{tc}" : $"item:{k.SquashItemId}")
            .Select(g => g.OrderByDescending(k => k.UltimaEjecucion ?? DateTime.MinValue).ThenByDescending(k => k.SquashItemId).First())
            .ToList();
}

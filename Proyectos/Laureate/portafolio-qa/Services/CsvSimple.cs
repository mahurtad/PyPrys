using System.Text;

namespace PortafolioQA.Services;

/// <summary>Lector CSV mínimo: detecta coma o punto y coma (Excel es-PE usa ';'), comillas dobles y saltos de línea dentro de comillas.</summary>
public static class CsvSimple
{
    public static List<string[]> Leer(string texto)
    {
        texto = texto.TrimStart('﻿');
        var primeraLinea = texto.Split('\n', 2)[0];
        var delimitador = primeraLinea.Count(c => c == ';') >= primeraLinea.Count(c => c == ',') && primeraLinea.Contains(';') ? ';' : ',';

        var filas = new List<string[]>();
        var fila = new List<string>();
        var campo = new StringBuilder();
        var entreComillas = false;

        void CerrarCampo() { fila.Add(campo.ToString().Trim()); campo.Clear(); }
        void CerrarFila()
        {
            CerrarCampo();
            if (fila.Any(c => c.Length > 0)) filas.Add(fila.ToArray());
            fila = new List<string>();
        }

        for (var i = 0; i < texto.Length; i++)
        {
            var c = texto[i];
            if (entreComillas)
            {
                if (c == '"' && i + 1 < texto.Length && texto[i + 1] == '"') { campo.Append('"'); i++; }
                else if (c == '"') entreComillas = false;
                else campo.Append(c);
            }
            else if (c == '"') entreComillas = true;
            else if (c == delimitador) CerrarCampo();
            else if (c == '\n') CerrarFila();
            else if (c != '\r') campo.Append(c);
        }
        if (campo.Length > 0 || fila.Count > 0) CerrarFila();
        return filas;
    }
}

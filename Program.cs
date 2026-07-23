using System.Globalization;
using System.Text;
using System.Text.RegularExpressions;
using PdfSharp.Pdf.IO;
using PdfSharpDocument = PdfSharp.Pdf.PdfDocument;
using PdfPigDocument = UglyToad.PdfPig.PdfDocument;

// ============================================================================
//  PDF Splitter por palabras clave
//  ------------------------------------------------------------------------
//  Recorre cada PDF de la carpeta "Entrada", detecta las paginas que
//  contienen alguna de las frases clave (busqueda sin distinguir mayusculas
//  ni acentos) y las trata como "inicio de documento". Genera un PDF nuevo
//  por cada rango [inicio -> siguiente inicio - 1] dentro de "Splits".
// ============================================================================

// --- Rutas base: Documentos\SplitPDF\{Entrada, Splits} ----------------------
string documentos = Environment.GetFolderPath(Environment.SpecialFolder.MyDocuments);
string baseDir = Path.Combine(documentos, "SplitPDF");
string entradaDir = Path.Combine(baseDir, "Entrada");
string splitsDir = Path.Combine(baseDir, "Splits");
string keywordsFile = Path.Combine(baseDir, "keywords.txt");

Directory.CreateDirectory(entradaDir);
Directory.CreateDirectory(splitsDir);

Console.OutputEncoding = Encoding.UTF8;
Console.WriteLine("== PDF Splitter por palabras clave ==");
Console.WriteLine($"Entrada : {entradaDir}");
Console.WriteLine($"Salida  : {splitsDir}");
Console.WriteLine();

// --- Cargar frases clave (se crea el archivo con ejemplos si no existe) -----
if (!File.Exists(keywordsFile))
{
    /*
     * "# Una frase clave por linea. Lineas que empiezan con # se ignoran.",
        "# La busqueda ignora mayusculas/minusculas y acentos.",
        "Estado de Cuenta",
        "Resumen de Movimientos",
     */
    File.WriteAllLines(keywordsFile, new[]
    {
        "Estado de Cuenta"
    }, new UTF8Encoding(false));
    Console.WriteLine($"Se creo el archivo de frases clave: {keywordsFile}");
    Console.WriteLine("Editalo con tus frases y vuelve a ejecutar.");
}

List<string> keywords = File.ReadAllLines(keywordsFile)
    .Select(l => l.Trim())
    .Where(l => l.Length > 0 && !l.StartsWith('#'))
    .ToList();

if (keywords.Count == 0)
{
    Console.WriteLine("No hay frases clave definidas en keywords.txt. Nada que hacer.");
    return;
}

Console.WriteLine("Frases clave:");
foreach (var k in keywords) Console.WriteLine($"  - {k}");
Console.WriteLine();

// Versiones normalizadas de las frases clave (una sola vez)
var keywordsNorm = keywords.Select(Normalize).ToList();

// --- Procesar cada PDF de la carpeta de entrada -----------------------------
var pdfs = Directory.GetFiles(entradaDir, "*.pdf", SearchOption.TopDirectoryOnly)
    .OrderBy(p => p, StringComparer.OrdinalIgnoreCase)
    .ToList();

if (pdfs.Count == 0)
{
    Console.WriteLine($"No hay PDFs en la carpeta de entrada.");
    Console.WriteLine($"Copia tus archivos en: {entradaDir}");
    return;
}

int totalGenerados = 0;
foreach (var pdfPath in pdfs)
{
    try
    {
        totalGenerados += ProcesarPdf(pdfPath, splitsDir, keywordsNorm);
    }
    catch (Exception ex)
    {
        Console.WriteLine($"  [ERROR] {Path.GetFileName(pdfPath)}: {ex.Message}");
    }
    Console.WriteLine();
}

Console.WriteLine($"Listo. Se generaron {totalGenerados} archivo(s) en {splitsDir}");


// ============================================================================
//  Funciones
// ============================================================================

int ProcesarPdf(string pdfPath, string outputRoot, List<string> palabras)
{
    string nombre = Path.GetFileNameWithoutExtension(pdfPath);
    Console.WriteLine($"> {Path.GetFileName(pdfPath)}");

    // 1-4) Extraer texto de cada pagina y detectar paginas de inicio
    string[] textoPorPagina;
    using (var doc = PdfPigDocument.Open(pdfPath))
    {
        textoPorPagina = doc.GetPages().Select(p => p.Text ?? string.Empty).ToArray();
    }

    int totalPaginas = textoPorPagina.Length;
    var startPages = new List<int>(); // indices base 0
    for (int i = 0; i < totalPaginas; i++)
    {
        string norm = Normalize(textoPorPagina[i]);
        if (palabras.Any(k => norm.Contains(k)))
            startPages.Add(i);
    }

    Console.WriteLine($"  Paginas: {totalPaginas} | Inicios detectados: {startPages.Count}");

    if (startPages.Count == 0)
    {
        Console.WriteLine("  No se encontro ninguna frase clave. Se omite este archivo.");
        return 0;
    }

    // Aviso si hay paginas antes del primer inicio (quedarian fuera de todo rango)
    if (startPages[0] > 0)
        Console.WriteLine($"  [Aviso] Las paginas 1-{startPages[0]} van antes del primer inicio y no se exportan.");

    // 6) Construir rangos y 7) exportar
    string outputDir = Path.Combine(outputRoot, SanitizeFileName(nombre));
    Directory.CreateDirectory(outputDir);

    int generados = 0;
    for (int i = 0; i < startPages.Count; i++)
    {
        int start = startPages[i];
        int end = (i < startPages.Count - 1) ? startPages[i + 1] - 1 : totalPaginas - 1;

        // 8) Nombre a partir de la primera pagina del rango (cuenta/mes/anio)
        string baseNombre = ConstruirNombre(textoPorPagina[start], i + 1);
        string destino = RutaUnica(outputDir, baseNombre);

        ExportarPaginas(pdfPath, start, end, destino);
        Console.WriteLine($"  [{i + 1}] paginas {start + 1}-{end + 1}  ->  {Path.GetFileName(destino)}");
        generados++;
    }

    return generados;
}

// Copia las paginas [start..end] (indices base 0, inclusivo) a un PDF nuevo
void ExportarPaginas(string origen, int start, int end, string destino)
{
    using var input = PdfReader.Open(origen, PdfDocumentOpenMode.Import);
    using var output = new PdfSharpDocument();
    for (int p = start; p <= end && p < input.PageCount; p++)
        output.AddPage(input.Pages[p]);
    output.Save(destino);
}

// Intenta armar un nombre con numero de cuenta / mes / anio de la portada.
// Si no encuentra nada, usa doc_NN.
string ConstruirNombre(string textoPagina, int indice)
{
    var partes = new List<string>();

    string? cuenta = ExtraerCuenta(textoPagina);
    if (cuenta != null) partes.Add("cuenta_" + cuenta);

    var (mes, anio) = ExtraerMesAnio(textoPagina);
    if (mes != null) partes.Add(mes);
    if (anio != null) partes.Add(anio);

    string nombre = partes.Count > 0
        ? string.Join("_", partes)
        : $"doc_{indice:00}";

    return SanitizeFileName(nombre) + ".pdf";
}

string? ExtraerCuenta(string texto)
{
    // Busca "cuenta / contrato / cliente / clabe" seguido de una secuencia de digitos
    var m = Regex.Match(texto,
        @"(?:no\.?\s*de\s*cuenta|n[uú]mero\s*de\s*cuenta|cuenta|contrato|clabe|cliente)\s*[:#\-]?\s*([0-9][0-9\-\s]{4,}[0-9])",
        RegexOptions.IgnoreCase);
    if (!m.Success) return null;
    string digitos = Regex.Replace(m.Groups[1].Value, @"[^0-9]", "");
    return digitos.Length >= 5 ? digitos : null;
}

(string? mes, string? anio) ExtraerMesAnio(string texto)
{
    string[] meses = { "enero", "febrero", "marzo", "abril", "mayo", "junio",
                       "julio", "agosto", "septiembre", "octubre", "noviembre", "diciembre" };
    string norm = Normalize(texto);

    string? mes = null;
    for (int i = 0; i < meses.Length; i++)
    {
        if (norm.Contains(meses[i]))
        {
            mes = $"{i + 1:00}_{meses[i]}";
            break;
        }
    }

    var my = Regex.Match(texto, @"\b(19|20)\d{2}\b");
    string? anio = my.Success ? my.Value : null;

    return (mes, anio);
}

// --- Utilidades -------------------------------------------------------------

// Minusculas, sin acentos y con espacios colapsados (para busqueda robusta)
string Normalize(string s)
{
    if (string.IsNullOrEmpty(s)) return string.Empty;
    string formD = s.ToLowerInvariant().Normalize(NormalizationForm.FormD);
    var sb = new StringBuilder(formD.Length);
    foreach (char c in formD)
    {
        if (CharUnicodeInfo.GetUnicodeCategory(c) != UnicodeCategory.NonSpacingMark)
            sb.Append(c);
    }
    string sinAcentos = sb.ToString().Normalize(NormalizationForm.FormC);
    return Regex.Replace(sinAcentos, @"\s+", " ").Trim();
}

string SanitizeFileName(string name)
{
    foreach (char c in Path.GetInvalidFileNameChars())
        name = name.Replace(c, '_');
    return name.Trim();
}

// Evita sobrescribir: agrega _2, _3, ... si ya existe
string RutaUnica(string dir, string fileName)
{
    string ruta = Path.Combine(dir, fileName);
    if (!File.Exists(ruta)) return ruta;

    string baseName = Path.GetFileNameWithoutExtension(fileName);
    string ext = Path.GetExtension(fileName);
    int n = 2;
    do
    {
        ruta = Path.Combine(dir, $"{baseName}_{n}{ext}");
        n++;
    } while (File.Exists(ruta));
    return ruta;
}

// ============================================================================
// MetadataExtractorBase.cs — Shared Infrastructure
// ============================================================================
// Contains the proven helper methods that BOTH V1 and V2 need:
//   - GetAllObjectNames()  — the model-iterating aggregator from your working code
//   - Console helpers       — colored output
//   - IO helpers            — StreamWriter creation, markdown cleaning
//   - Type helpers          — EDT/Field type resolution
// ============================================================================

using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using Microsoft.Dynamics.AX.Metadata.MetaModel;
using Microsoft.Dynamics.AX.Metadata.Providers;

namespace D365FOMetadataExtractor
{
    /// <summary>
    /// Base class providing shared infrastructure for all extractor versions.
    /// </summary>
    public abstract class MetadataExtractorBase
    {
        protected IMetadataProvider MetadataProvider { get; private set; }
        protected List<string> ModelNames { get; private set; }
        protected string OutputDirectory { get; private set; }
        protected HashSet<ExtractionCategory> SelectedCategories { get; private set; }
        protected Dictionary<string, int> ObjectCounts { get; } = new Dictionary<string, int>();

        /// <summary>
        /// Called by Program.Main() via the interface. Sets up shared state then delegates to RunExtraction().
        /// </summary>
        public void Execute(IMetadataProvider metadataProvider, List<string> modelNames, string outputDirectory, HashSet<ExtractionCategory> selectedCategories)
        {
            MetadataProvider = metadataProvider;
            ModelNames = modelNames;
            OutputDirectory = outputDirectory;
            SelectedCategories = selectedCategories ?? new HashSet<ExtractionCategory>();

            Directory.CreateDirectory(outputDirectory);
            ObjectCounts.Clear();

            RunExtraction();

            PrintSummary();
        }

        /// <summary>
        /// Checks if a category should be extracted based on user selection.
        /// </summary>
        protected bool ShouldExtract(ExtractionCategory category)
        {
            return SelectedCategories.Contains(category);
        }

        /// <summary>
        /// Subclasses implement this with their specific extraction logic.
        /// </summary>
        protected abstract void RunExtraction();

        // ═══════════════════════════════════════════════════════════════
        //  CORE: Model-iterating aggregator (from your working V1 code)
        // ═══════════════════════════════════════════════════════════════

        /// <summary>
        /// Iterates every discovered model and aggregates object names.
        /// HashSet for dedup — same object can appear across models.
        /// Try/catch per model so restricted models don't kill the run.
        /// THIS IS THE PROVEN PATTERN FROM THE WORKING V1 CODE.
        /// </summary>
        protected List<string> GetAllObjectNames(Func<string, IEnumerable<string>> listFunc)
        {
            var all = new HashSet<string>();
            int modelCount = 0;
            int successCount = 0;
            int emptyCount = 0;
            int errorCount = 0;

            foreach (var model in ModelNames)
            {
                modelCount++;
                try
                {
                    var objects = listFunc(model).ToList();
                    int objectCount = objects.Count;

                    if (objectCount > 0)
                    {
                        Console.WriteLine($"  [{modelCount}/{ModelNames.Count}] {model,-40} : {objectCount,5} objects");
                        foreach (var obj in objects)
                        {
                            all.Add(obj);
                        }
                        successCount++;
                    }
                    else
                    {
                        emptyCount++;
                    }
                }
                catch (Exception ex)
                {
                    // Some models may be empty or restricted — log and skip
                    Console.ForegroundColor = ConsoleColor.DarkYellow;
                    Console.WriteLine($"  [{modelCount}/{ModelNames.Count}] {model,-40} : ERROR - {ex.Message}");
                    Console.ResetColor();
                    errorCount++;
                }
            }

            Console.WriteLine($"\n  Summary: {successCount} models with objects, {emptyCount} empty, {errorCount} errors");
            Console.WriteLine($"  Total unique objects found: {all.Count:N0}");

            return all.OrderBy(x => x).ToList();
        }

        // ═══════════════════════════════════════════════════════════════
        //  IO HELPERS
        // ═══════════════════════════════════════════════════════════════

        /// <summary>
        /// Creates a StreamWriter for a file inside the output directory.
        /// </summary>
        protected StreamWriter CreateWriter(string fileName)
        {
            string path = Path.Combine(OutputDirectory, fileName);
            return new StreamWriter(path, false, Encoding.UTF8);
        }

        /// <summary>
        /// Writes a simple markdown list section and flushes.
        /// Used by V1's simple format and V2's combined skill file.
        /// </summary>
        protected void WriteSection(StreamWriter writer, string sectionName, IEnumerable<string> items)
        {
            Console.WriteLine($"   Writing: {sectionName}");
            writer.WriteLine($"## {sectionName}");
            foreach (var item in items)
            {
                writer.WriteLine($"- {item}");
            }
            writer.WriteLine();
            writer.Flush();
        }

        /// <summary>
        /// Cleans text for safe markdown rendering (strips pipes, newlines).
        /// </summary>
        protected string Clean(string text)
        {
            if (string.IsNullOrEmpty(text)) return "";
            return text
                .Replace("|", "\\|")
                .Replace("\r\n", " ")
                .Replace("\n", " ")
                .Replace("\r", " ")
                .Trim();
        }

        /// <summary>
        /// Records an object count for the final summary.
        /// </summary>
        protected void RecordCount(string category, int count)
        {
            ObjectCounts[category] = count;
            WriteSuccess($"   -> {count:N0} {category}");
        }

        // ═══════════════════════════════════════════════════════════════
        //  TYPE HELPERS — EDT, Table Field, Map Field type resolution
        // ═══════════════════════════════════════════════════════════════

        protected string GetEdtBaseType(AxEdt edt)
        {
            if (edt is AxEdtString)    return "String";
            if (edt is AxEdtInt)       return "Int";
            if (edt is AxEdtInt64)     return "Int64";
            if (edt is AxEdtReal)      return "Real";
            if (edt is AxEdtDate)      return "Date";
            if (edt is AxEdtUtcDateTime) return "UtcDateTime";
            if (edt is AxEdtGuid)      return "Guid";
            if (edt is AxEdtContainer) return "Container";
            if (edt is AxEdtEnum edtE) return $"Enum ({edtE.EnumType})";
            return edt.GetType().Name.Replace("AxEdt", "");
        }

        protected string GetFieldType(AxTableField field)
        {
            if (field is AxTableFieldString)      return "String";
            if (field is AxTableFieldInt)          return "Int";
            if (field is AxTableFieldInt64)        return "Int64";
            if (field is AxTableFieldReal)         return "Real";
            if (field is AxTableFieldDate)         return "Date";
            if (field is AxTableFieldUtcDateTime)  return "UtcDateTime";
            if (field is AxTableFieldGuid)         return "Guid";
            if (field is AxTableFieldContainer)    return "Container";
            if (field is AxTableFieldEnum)         return "Enum";
            if (field is AxTableFieldTime)         return "Time";
            return field.GetType().Name.Replace("AxTableField", "");
        }

        protected string GetMapFieldType(object field)
        {
            if (field is AxMapFieldString)      return "String";
            if (field is AxMapFieldInt)          return "Int";
            if (field is AxMapFieldInt64)        return "Int64";
            if (field is AxMapFieldReal)         return "Real";
            if (field is AxMapFieldDate)         return "Date";
            if (field is AxMapFieldUtcDateTime)  return "UtcDateTime";
            if (field is AxMapFieldGuid)         return "Guid";
            if (field is AxMapFieldContainer)    return "Container";
            if (field is AxMapFieldEnum)         return "Enum";
            return field.GetType().Name.Replace("AxMapField", "");
        }

        protected string GetFieldEdtOrEnum(AxTableField field)
        {
            if (!string.IsNullOrWhiteSpace(field.ExtendedDataType))
                return field.ExtendedDataType;
            if (field is AxTableFieldEnum enumField && !string.IsNullOrWhiteSpace(enumField.EnumType))
                return enumField.EnumType;
            return "";
        }

        // ═══════════════════════════════════════════════════════════════
        //  CONSOLE HELPERS
        // ═══════════════════════════════════════════════════════════════

        protected void WriteStep(string msg)
        {
            Console.ForegroundColor = ConsoleColor.Yellow;
            Console.WriteLine($"\n> {msg}");
            Console.ResetColor();
        }

        protected void WriteSuccess(string msg)
        {
            Console.ForegroundColor = ConsoleColor.Green;
            Console.WriteLine(msg);
            Console.ResetColor();
        }

        protected void WriteError(string msg)
        {
            Console.ForegroundColor = ConsoleColor.Red;
            Console.WriteLine($"  ERROR: {msg}");
            Console.ResetColor();
        }

        private void PrintSummary()
        {
            Console.WriteLine();
            Console.WriteLine("Object counts:");
            long total = 0;
            foreach (var kvp in ObjectCounts.OrderByDescending(x => x.Value))
            {
                Console.WriteLine($"  {kvp.Key,-35} {kvp.Value,8:N0}");
                total += kvp.Value;
            }
            Console.WriteLine($"  {"TOTAL",-35} {total,8:N0}");
        }
    }
}

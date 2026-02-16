// ============================================================================
// DetailedMetadataExtractor.cs — Version 2 (Enhanced Edition)
// ============================================================================
// DEFENSIVE APPROACH: All property access on metadata objects uses TryGetProp()
// which reads via reflection. If a property doesn't exist in this DLL version,
// it returns "" instead of a compile error. Only .Name is accessed directly
// since every metadata object guarantees it.
// ============================================================================

using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.IO;
using System.Linq;
using System.Reflection;
using System.Text;
using Microsoft.Dynamics.AX.Metadata.MetaModel;

namespace D365FOMetadataExtractor
{
    public class DetailedMetadataExtractor : MetadataExtractorBase, IMetadataExtractor
    {
        public string VersionName => "V2 — Detailed (Fields, Indexes, Relations) - Combined Output";

        private StreamWriter _combinedWriter;

        protected override void RunExtraction()
        {
            string outputFile = Path.Combine(OutputDirectory, "D365FO_Metadata.md");

            using (_combinedWriter = new StreamWriter(outputFile, false, Encoding.UTF8))
            {
                // Write header
                _combinedWriter.WriteLine("# D365FO Metadata Extraction");
                _combinedWriter.WriteLine();
                _combinedWriter.WriteLine($"**Extracted:** {DateTime.Now:yyyy-MM-dd HH:mm:ss}");
                _combinedWriter.WriteLine($"**Models:** {ModelNames.Count}");
                _combinedWriter.WriteLine();
                _combinedWriter.WriteLine("---");
                _combinedWriter.WriteLine();

                // Extract selected categories
                if (ShouldExtract(ExtractionCategory.Enums)) ExtractEnums();
                if (ShouldExtract(ExtractionCategory.Edts)) ExtractEdts();
                if (ShouldExtract(ExtractionCategory.Tables)) ExtractTables();
                if (ShouldExtract(ExtractionCategory.Views)) ExtractViews();
                if (ShouldExtract(ExtractionCategory.DataEntities)) ExtractDataEntities();
                if (ShouldExtract(ExtractionCategory.Classes)) ExtractClasses();
                if (ShouldExtract(ExtractionCategory.Forms)) ExtractForms();
                if (ShouldExtract(ExtractionCategory.MenuItems)) ExtractMenuItems();
                if (ShouldExtract(ExtractionCategory.Queries)) ExtractQueries();
                if (ShouldExtract(ExtractionCategory.Services)) ExtractServices();
                if (ShouldExtract(ExtractionCategory.Maps)) ExtractMaps();
                if (ShouldExtract(ExtractionCategory.SecurityRoles) ||
                    ShouldExtract(ExtractionCategory.SecurityDuties) ||
                    ShouldExtract(ExtractionCategory.SecurityPrivileges))
                {
                    ExtractSecurityObjects();
                }
                if (ShouldExtract(ExtractionCategory.CompositeDataEntities) ||
                    ShouldExtract(ExtractionCategory.AggregateDataEntities))
                {
                    ExtractCompositeAndAggregate();
                }

                _combinedWriter.Flush();
            }

            WriteSuccess($"\n   Output file: {outputFile}");
        }

        // ═══════════════════════════════════════════════════════════════
        //  SAFE PROPERTY ACCESS — reflection-based, never throws
        // ═══════════════════════════════════════════════════════════════

        /// <summary>
        /// Safely reads a property by name via reflection.
        /// Returns the string value, or defaultVal if the property doesn't exist.
        /// This eliminates ALL compile errors from missing properties across DLL versions.
        /// </summary>
        private string TryGetProp(object obj, string propertyName, string defaultVal = "")
        {
            if (obj == null) return defaultVal;
            try
            {
                PropertyInfo prop = obj.GetType().GetProperty(propertyName,
                    BindingFlags.Public | BindingFlags.Instance);
                if (prop == null) return defaultVal;

                object val = prop.GetValue(obj);
                if (val == null) return defaultVal;

                return val.ToString();
            }
            catch
            {
                return defaultVal;
            }
        }

        /// <summary>
        /// Safely reads a property and returns the raw object (for collections).
        /// Returns null if the property doesn't exist.
        /// </summary>
        private object TryGetObj(object obj, string propertyName)
        {
            if (obj == null) return null;
            try
            {
                PropertyInfo prop = obj.GetType().GetProperty(propertyName,
                    BindingFlags.Public | BindingFlags.Instance);
                return prop?.GetValue(obj);
            }
            catch
            {
                return null;
            }
        }

        // ═══════════════════════════════════════════════════════════════
        //  ENUMS
        // ═══════════════════════════════════════════════════════════════

        private void ExtractEnums()
        {
            WriteStep("Extracting Base Enums (with values)...");
            var names = GetAllObjectNames(m => MetadataProvider.Enums.ListObjects(m));

            _combinedWriter.WriteLine("# Base Enums");
            _combinedWriter.WriteLine();
            _combinedWriter.WriteLine("Each enum lists its symbolic values. Use these exact names in X++ code.");
            _combinedWriter.WriteLine();

            int count = 0;
            foreach (var name in names)
            {
                try
                {
                    var axEnum = MetadataProvider.Enums.Read(name);
                    if (axEnum == null) continue;

                    _combinedWriter.WriteLine($"## {axEnum.Name}");

                    string label = TryGetProp(axEnum, "Label");
                    if (!string.IsNullOrWhiteSpace(label))
                        _combinedWriter.WriteLine($"Label: {label}");

                    if (axEnum.EnumValues != null && axEnum.EnumValues.Any())
                    {
                        _combinedWriter.WriteLine("| Value | Name |");
                        _combinedWriter.WriteLine("|-------|------|");
                        foreach (var val in axEnum.EnumValues)
                            _combinedWriter.WriteLine($"| {val.Value} | {val.Name} |");
                    }
                    _combinedWriter.WriteLine();
                    _combinedWriter.Flush();
                    count++;
                }
                catch (Exception ex)
                {
                    Debug.WriteLine($"Enum [{name}]: {ex.Message}");
                }
            }
            _combinedWriter.WriteLine("---");
            _combinedWriter.WriteLine();
            RecordCount("Base Enums", count);
        }

        // ═══════════════════════════════════════════════════════════════
        //  EDTs
        // ═══════════════════════════════════════════════════════════════

        private void ExtractEdts()
        {
            WriteStep("Extracting Extended Data Types...");
            var names = GetAllObjectNames(m => MetadataProvider.Edts.ListObjects(m));

            _combinedWriter.WriteLine("# Extended Data Types (EDTs)");
            _combinedWriter.WriteLine();
            _combinedWriter.WriteLine("| EDT Name | Extends | Base Type | String Length | Label |");
            _combinedWriter.WriteLine("|----------|---------|-----------|--------------|-------|");

            int count = 0;
            foreach (var name in names)
            {
                try
                {
                    var edt = MetadataProvider.Edts.Read(name);
                    if (edt == null) continue;

                    string extends  = TryGetProp(edt, "Extends");
                    string baseType = GetEdtBaseType(edt);
                    string strLen   = "";
                    string label    = Clean(TryGetProp(edt, "Label"));

                    if (edt is AxEdtString edtStr)
                    {
                        string size = TryGetProp(edtStr, "StringSize");
                        if (!string.IsNullOrWhiteSpace(size) && size != "0")
                            strLen = size;
                    }

                    _combinedWriter.WriteLine($"| {edt.Name} | {extends} | {baseType} | {strLen} | {label} |");
                    count++;
                }
                catch (Exception ex)
                {
                    Debug.WriteLine($"EDT [{name}]: {ex.Message}");
                }
            }
            _combinedWriter.WriteLine();
            _combinedWriter.WriteLine("---");
            _combinedWriter.WriteLine();
            _combinedWriter.Flush();
            RecordCount("Extended Data Types", count);
        }

        // ═══════════════════════════════════════════════════════════════
        //  TABLES — the most critical extraction
        // ═══════════════════════════════════════════════════════════════

        private void ExtractTables()
        {
            WriteStep("Extracting Tables (fields, indexes, relations — may take several minutes)...");
            var names = GetAllObjectNames(m => MetadataProvider.Tables.ListObjects(m));
            int total = names.Count;

            _combinedWriter.WriteLine("# Tables");
            _combinedWriter.WriteLine();
            _combinedWriter.WriteLine("Each table lists fields with EDT/Enum type, indexes, and relations.");
            _combinedWriter.WriteLine();

            int count = 0;
            int failedCount = 0;
            foreach (var name in names)
                {
                    try
                    {
                        var table = MetadataProvider.Tables.Read(name);
                        if (table == null)
                        {
                            Console.ForegroundColor = ConsoleColor.DarkYellow;
                            Console.WriteLine($"  WARNING: Table '{name}' returned null from Read()");
                            Console.ResetColor();
                            failedCount++;
                            continue;
                        }

                        _combinedWriter.WriteLine($"## {table.Name}");

                        // Safe property reads for optional table metadata
                        string label = TryGetProp(table, "Label");
                        if (!string.IsNullOrWhiteSpace(label))
                            _combinedWriter.WriteLine($"Label: {label}");

                        var props = new List<string>();
                        string tableGroup  = TryGetProp(table, "TableGroup");
                        string primaryIdx  = TryGetProp(table, "PrimaryIndex");
                        string clusterIdx  = TryGetProp(table, "ClusterIndex");
                        string extends     = TryGetProp(table, "Extends");

                        if (!string.IsNullOrWhiteSpace(tableGroup) && tableGroup != "Miscellaneous")
                            props.Add($"Group: {tableGroup}");
                        if (!string.IsNullOrWhiteSpace(primaryIdx))
                            props.Add($"PrimaryIndex: {primaryIdx}");
                        if (!string.IsNullOrWhiteSpace(clusterIdx))
                            props.Add($"ClusterIndex: {clusterIdx}");
                        if (!string.IsNullOrWhiteSpace(extends))
                            props.Add($"Extends: {extends}");
                        if (props.Any())
                            _combinedWriter.WriteLine(string.Join(" | ", props));

                        // ── Fields (AxTable.Fields is reliable across versions) ──
                        if (table.Fields != null && table.Fields.Any())
                        {
                            _combinedWriter.WriteLine("| Field | Type | EDT/Enum | Mandatory |");
                            _combinedWriter.WriteLine("|-------|------|----------|-----------|");
                            foreach (var field in table.Fields)
                            {
                                string fType    = GetFieldType(field);
                                string fEdtEnum = GetFieldEdtOrEnum(field);
                                string fMand    = TryGetProp(field, "Mandatory");
                                string mandDisplay = (fMand == "Yes" || fMand == "1") ? "Yes" : "";

                                _combinedWriter.WriteLine($"| {field.Name} | {fType} | {fEdtEnum} | {mandDisplay} |");
                            }
                        }

                        // ── Field Groups ──
                        if (table.FieldGroups != null && table.FieldGroups.Any())
                        {
                            _combinedWriter.WriteLine();
                            _combinedWriter.WriteLine("Field Groups:");
                            foreach (var fg in table.FieldGroups)
                            {
                                try
                                {
                                    // fg.Fields contains AxTableFieldGroupField items
                                    var fgFieldsObj = TryGetObj(fg, "Fields") as System.Collections.IEnumerable;
                                    if (fgFieldsObj != null)
                                    {
                                        var fieldNames = new List<string>();
                                        foreach (var fgf in fgFieldsObj)
                                        {
                                            string df = TryGetProp(fgf, "DataField");
                                            if (!string.IsNullOrWhiteSpace(df))
                                                fieldNames.Add(df);
                                        }
                                        if (fieldNames.Any())
                                            _combinedWriter.WriteLine($"- {fg.Name}: {string.Join(", ", fieldNames)}");
                                    }
                                }
                                catch { }
                            }
                        }

                        // ── Indexes ──
                        if (table.Indexes != null && table.Indexes.Any())
                        {
                            _combinedWriter.WriteLine();
                            _combinedWriter.WriteLine("Indexes:");
                            foreach (var idx in table.Indexes)
                            {
                                try
                                {
                                    string allowDup = TryGetProp(idx, "AllowDuplicates");
                                    string unique = (allowDup == "No" || allowDup == "0") ? " (Unique)" : "";

                                    var idxFieldsObj = TryGetObj(idx, "Fields") as System.Collections.IEnumerable;
                                    var idxFieldNames = new List<string>();
                                    if (idxFieldsObj != null)
                                    {
                                        foreach (var idxF in idxFieldsObj)
                                        {
                                            string df = TryGetProp(idxF, "DataField");
                                            if (!string.IsNullOrWhiteSpace(df))
                                                idxFieldNames.Add(df);
                                        }
                                    }

                                    _combinedWriter.WriteLine($"- {idx.Name}{unique}: {string.Join(", ", idxFieldNames)}");
                                }
                                catch { }
                            }
                        }

                        // ── Relations ──
                        if (table.Relations != null && table.Relations.Any())
                        {
                            _combinedWriter.WriteLine();
                            _combinedWriter.WriteLine("Relations:");
                            foreach (var rel in table.Relations)
                            {
                                string relatedTable = TryGetProp(rel, "RelatedTable");
                                _combinedWriter.WriteLine($"- {rel.Name} -> {relatedTable}");
                            }
                        }

                        _combinedWriter.WriteLine();
                        _combinedWriter.Flush();
                        count++;

                        if (count % 500 == 0)
                            Console.WriteLine($"   Tables: {count:N0} / {total:N0}");
                    }
                    catch (Exception ex)
                    {
                        Console.ForegroundColor = ConsoleColor.DarkYellow;
                        Console.WriteLine($"  ERROR reading table '{name}': {ex.Message}");
                        Console.ResetColor();
                        Debug.WriteLine($"Table [{name}]: {ex.Message}");
                        failedCount++;
                    }
                }
                _combinedWriter.WriteLine("---");
                _combinedWriter.WriteLine();
                RecordCount("Tables", count);
                if (failedCount > 0)
                {
                    Console.ForegroundColor = ConsoleColor.Yellow;
                    Console.WriteLine($"   -> {failedCount:N0} tables failed to read (see warnings above)");
                    Console.ResetColor();
                }
        }

        // ═══════════════════════════════════════════════════════════════
        //  VIEWS
        // ═══════════════════════════════════════════════════════════════

        private void ExtractViews()
        {
            WriteStep("Extracting Views...");
            var names = GetAllObjectNames(m => MetadataProvider.Views.ListObjects(m));

            _combinedWriter.WriteLine("# Views");
            _combinedWriter.WriteLine();

            int count = 0;
            foreach (var name in names)
            {
                try
                {
                    var view = MetadataProvider.Views.Read(name);
                    if (view == null) continue;

                    _combinedWriter.WriteLine($"## {view.Name}");

                    string label = TryGetProp(view, "Label");
                    if (!string.IsNullOrWhiteSpace(label))
                        _combinedWriter.WriteLine($"Label: {label}");

                    // ViewMetadata.Fields — access safely
                    var viewMeta = TryGetObj(view, "ViewMetadata");
                    var fields = TryGetObj(viewMeta, "Fields") as System.Collections.IEnumerable;
                    if (fields != null)
                    {
                        _combinedWriter.WriteLine("| Field | Type |");
                        _combinedWriter.WriteLine("|-------|------|");
                        foreach (var field in fields)
                        {
                            string fName = TryGetProp(field, "Name");
                            string fType = field.GetType().Name.Replace("AxViewField", "");
                            _combinedWriter.WriteLine($"| {fName} | {fType} |");
                        }
                    }

                    _combinedWriter.WriteLine();
                    _combinedWriter.Flush();
                    count++;
                }
                catch (Exception ex)
                {
                    Debug.WriteLine($"View [{name}]: {ex.Message}");
                }
            }
            _combinedWriter.WriteLine("---");
            _combinedWriter.WriteLine();
            RecordCount("Views", count);
        }

        // ═══════════════════════════════════════════════════════════════
        //  DATA ENTITIES
        // ═══════════════════════════════════════════════════════════════

        private void ExtractDataEntities()
        {
            WriteStep("Extracting Data Entities...");
            var names = GetAllObjectNames(m => MetadataProvider.DataEntityViews.ListObjects(m));

            _combinedWriter.WriteLine("# Data Entities");
            _combinedWriter.WriteLine();

            int count = 0;
            foreach (var name in names)
            {
                try
                {
                    var entity = MetadataProvider.DataEntityViews.Read(name);
                    if (entity == null) continue;

                    _combinedWriter.WriteLine($"## {entity.Name}");

                    string label      = TryGetProp(entity, "Label");
                    string pubName    = TryGetProp(entity, "PublicEntityName");
                    string pubColl    = TryGetProp(entity, "PublicCollectionName");
                    string isPublic   = TryGetProp(entity, "IsPublic");

                    if (!string.IsNullOrWhiteSpace(label))     _combinedWriter.WriteLine($"Label: {label}");
                    if (!string.IsNullOrWhiteSpace(pubName))   _combinedWriter.WriteLine($"Public Name: {pubName}");
                    if (!string.IsNullOrWhiteSpace(pubColl))   _combinedWriter.WriteLine($"Public Collection: {pubColl}");
                    if (!string.IsNullOrWhiteSpace(isPublic))  _combinedWriter.WriteLine($"Public: {isPublic}");

                    var viewMeta = TryGetObj(entity, "ViewMetadata");
                    var fields = TryGetObj(viewMeta, "Fields") as System.Collections.IEnumerable;
                    if (fields != null)
                    {
                        _combinedWriter.WriteLine("| Field | Type |");
                        _combinedWriter.WriteLine("|-------|------|");
                        foreach (var field in fields)
                        {
                            string fName = TryGetProp(field, "Name");
                            string fType = field.GetType().Name
                                .Replace("AxDataEntityViewField", "")
                                .Replace("AxViewField", "");
                            _combinedWriter.WriteLine($"| {fName} | {fType} |");
                        }
                    }

                    _combinedWriter.WriteLine();
                    _combinedWriter.Flush();
                    count++;
                }
                catch (Exception ex)
                {
                    Debug.WriteLine($"DataEntity [{name}]: {ex.Message}");
                }
            }
            _combinedWriter.WriteLine("---");
            _combinedWriter.WriteLine();
            RecordCount("Data Entities", count);
        }

        // ═══════════════════════════════════════════════════════════════
        //  CLASSES
        // ═══════════════════════════════════════════════════════════════

        private void ExtractClasses()
        {
            WriteStep("Extracting Classes...");
            var names = GetAllObjectNames(m => MetadataProvider.Classes.ListObjects(m));

            var keyPrefixes = new[]
            {
                "NumberSeq", "DimensionDefaulting", "InventPosting",
                "LedgerVoucher", "SalesFormLetter", "PurchFormLetter",
                "InventMov", "InventUpd", "InventTrans",
                "TaxCalc", "MarkupCalc",
                "SysOperation", "SysDataEntity", "DMF",
                "BatchHeader", "RunBase",
                "Query", "SysQuery", "FormLetter",
                "SysLookup", "EventHandler"
            };

            _combinedWriter.WriteLine("# Classes");
            _combinedWriter.WriteLine();
            _combinedWriter.WriteLine("## Key Framework Classes (with methods)");
            _combinedWriter.WriteLine();

            int detailCount = 0;
            foreach (var name in names)
            {
                bool isKey = keyPrefixes.Any(p =>
                    name.StartsWith(p, StringComparison.OrdinalIgnoreCase));
                if (!isKey) continue;

                try
                {
                    var axClass = MetadataProvider.Classes.Read(name);
                    if (axClass == null) continue;

                    _combinedWriter.WriteLine($"### {axClass.Name}");

                    string extends = TryGetProp(axClass, "Extends");
                    if (!string.IsNullOrWhiteSpace(extends))
                        _combinedWriter.WriteLine($"Extends: {extends}");

                    if (axClass.Methods != null && axClass.Methods.Any())
                    {
                        _combinedWriter.WriteLine("Methods:");
                        foreach (var method in axClass.Methods.OrderBy(m => m.Name))
                        {
                            string isStatic = TryGetProp(method, "IsStatic");
                            string st = (isStatic == "True" || isStatic == "1") ? "static " : "";
                            _combinedWriter.WriteLine($"- {st}{method.Name}");
                        }
                    }
                    _combinedWriter.WriteLine();
                    _combinedWriter.Flush();
                    detailCount++;
                }
                catch (Exception ex)
                {
                    Debug.WriteLine($"Class detail [{name}]: {ex.Message}");
                }
            }

            _combinedWriter.WriteLine("## All Class Names");
            _combinedWriter.WriteLine();
            _combinedWriter.WriteLine("```");
            foreach (var name in names)
                _combinedWriter.WriteLine(name);
            _combinedWriter.WriteLine("```");
            _combinedWriter.WriteLine();
            _combinedWriter.WriteLine("---");
            _combinedWriter.WriteLine();
            _combinedWriter.Flush();

            RecordCount("Classes (total)", names.Count);
            RecordCount("Classes (detailed)", detailCount);
        }

        // ═══════════════════════════════════════════════════════════════
        //  FORMS — name only (AxForm may not have .Label)
        // ═══════════════════════════════════════════════════════════════

        private void ExtractForms()
        {
            WriteStep("Extracting Forms...");
            var names = GetAllObjectNames(m => MetadataProvider.Forms.ListObjects(m));

            _combinedWriter.WriteLine("# Forms");
            _combinedWriter.WriteLine();
            _combinedWriter.WriteLine("Use these exact names with `new Args()` name or `MenuFunction`.");
            _combinedWriter.WriteLine();

            int count = 0;
            foreach (var name in names)
            {
                try
                {
                    var form = MetadataProvider.Forms.Read(name);
                    if (form == null) continue;

                    // Use TryGetProp — AxForm.Label doesn't exist in all DLL versions
                    string label = TryGetProp(form, "Label");
                    string display = !string.IsNullOrWhiteSpace(label) ? $" — {Clean(label)}" : "";
                    _combinedWriter.WriteLine($"- {form.Name}{display}");
                    count++;
                }
                catch { }
            }
            _combinedWriter.WriteLine();
            _combinedWriter.WriteLine("---");
            _combinedWriter.WriteLine();
            _combinedWriter.Flush();
            RecordCount("Forms", count);
        }

        // ═══════════════════════════════════════════════════════════════
        //  MENU ITEMS
        // ═══════════════════════════════════════════════════════════════

        private void ExtractMenuItems()
        {
            WriteStep("Extracting Menu Items...");

            _combinedWriter.WriteLine("# Menu Items");
            _combinedWriter.WriteLine();

            _combinedWriter.WriteLine("## Display Menu Items");
            int dCount = ExtractMenuItemCategory(_combinedWriter, m => MetadataProvider.MenuItemDisplays.ListObjects(m),
                name => MetadataProvider.MenuItemDisplays.Read(name));

            _combinedWriter.WriteLine();
            _combinedWriter.WriteLine("## Action Menu Items");
            int aCount = ExtractMenuItemCategory(_combinedWriter, m => MetadataProvider.MenuItemActions.ListObjects(m),
                name => MetadataProvider.MenuItemActions.Read(name));

            _combinedWriter.WriteLine();
            _combinedWriter.WriteLine("## Output Menu Items");
            int oCount = ExtractMenuItemCategory(_combinedWriter, m => MetadataProvider.MenuItemOutputs.ListObjects(m),
                name => MetadataProvider.MenuItemOutputs.Read(name));

            _combinedWriter.WriteLine("---");
            _combinedWriter.WriteLine();
            _combinedWriter.Flush();
            RecordCount("Menu Items (Display)", dCount);
            RecordCount("Menu Items (Action)", aCount);
            RecordCount("Menu Items (Output)", oCount);
        }

        private int ExtractMenuItemCategory(StreamWriter writer,
            Func<string, IEnumerable<string>> listFunc,
            Func<string, object> readFunc)
        {
            var names = GetAllObjectNames(listFunc);
            int count = 0;
            foreach (var name in names)
            {
                try
                {
                    var mi = readFunc(name);
                    if (mi == null) continue;
                    string miName = TryGetProp(mi, "Name");
                    string obj    = TryGetProp(mi, "Object");
                    string suffix = !string.IsNullOrWhiteSpace(obj) ? $" -> {obj}" : "";
                    writer.WriteLine($"- {miName}{suffix}");
                    count++;
                }
                catch { }
            }
            return count;
        }

        // ═══════════════════════════════════════════════════════════════
        //  QUERIES — name only (AxQuery may lack .Label, .DataSources)
        // ═══════════════════════════════════════════════════════════════

        private void ExtractQueries()
        {
            WriteStep("Extracting Queries...");
            var names = GetAllObjectNames(m => MetadataProvider.Queries.ListObjects(m));

            _combinedWriter.WriteLine("# Queries");
            _combinedWriter.WriteLine();

            int count = 0;
            foreach (var name in names)
            {
                try
                {
                    var query = MetadataProvider.Queries.Read(name);
                    if (query == null) continue;

                    _combinedWriter.WriteLine($"## {query.Name}");

                    // Safe — these properties may not exist
                    string label = TryGetProp(query, "Label");
                    if (!string.IsNullOrWhiteSpace(label))
                        _combinedWriter.WriteLine($"Label: {label}");

                    // DataSources may not exist on AxQuery in this version
                    var dataSources = TryGetObj(query, "DataSources") as System.Collections.IEnumerable;
                    if (dataSources != null)
                    {
                        _combinedWriter.WriteLine("Data Sources:");
                        foreach (var ds in dataSources)
                        {
                            string dsName  = TryGetProp(ds, "Name");
                            string dsTable = TryGetProp(ds, "Table");
                            _combinedWriter.WriteLine($"- {dsName} (Table: {dsTable})");
                        }
                    }

                    _combinedWriter.WriteLine();
                    _combinedWriter.Flush();
                    count++;
                }
                catch { }
            }
            _combinedWriter.WriteLine("---");
            _combinedWriter.WriteLine();
            RecordCount("Queries", count);
        }

        // ═══════════════════════════════════════════════════════════════
        //  SERVICES
        // ═══════════════════════════════════════════════════════════════

        private void ExtractServices()
        {
            WriteStep("Extracting Services...");
            var names = GetAllObjectNames(m => MetadataProvider.Services.ListObjects(m));

            _combinedWriter.WriteLine("# Services");
            _combinedWriter.WriteLine();

            int count = 0;
            foreach (var name in names)
            {
                try
                {
                    var svc = MetadataProvider.Services.Read(name);
                    if (svc == null) continue;

                    _combinedWriter.WriteLine($"## {svc.Name}");

                    string extName = TryGetProp(svc, "ExternalName");
                    string cls     = TryGetProp(svc, "Class");

                    if (!string.IsNullOrWhiteSpace(extName)) _combinedWriter.WriteLine($"External Name: {extName}");
                    if (!string.IsNullOrWhiteSpace(cls))     _combinedWriter.WriteLine($"Class: {cls}");

                    var ops = TryGetObj(svc, "ServiceOperations") as System.Collections.IEnumerable;
                    if (ops != null)
                    {
                        _combinedWriter.WriteLine("Operations:");
                        foreach (var op in ops)
                            _combinedWriter.WriteLine($"- {TryGetProp(op, "Name")}");
                    }

                    _combinedWriter.WriteLine();
                    _combinedWriter.Flush();
                    count++;
                }
                catch { }
            }
            _combinedWriter.WriteLine("---");
            _combinedWriter.WriteLine();
            RecordCount("Services", count);
        }

        // ═══════════════════════════════════════════════════════════════
        //  MAPS
        // ═══════════════════════════════════════════════════════════════

        private void ExtractMaps()
        {
            WriteStep("Extracting Maps...");
            var names = GetAllObjectNames(m => MetadataProvider.Maps.ListObjects(m));

            _combinedWriter.WriteLine("# Maps");
            _combinedWriter.WriteLine();

            int count = 0;
            foreach (var name in names)
            {
                try
                {
                    var map = MetadataProvider.Maps.Read(name);
                    if (map == null) continue;

                    _combinedWriter.WriteLine($"## {map.Name}");

                    var fields = TryGetObj(map, "Fields") as System.Collections.IEnumerable;
                    if (fields != null)
                    {
                        _combinedWriter.WriteLine("| Field | Type |");
                        _combinedWriter.WriteLine("|-------|------|");
                        foreach (var field in fields)
                            _combinedWriter.WriteLine($"| {TryGetProp(field, "Name")} | {GetMapFieldType(field)} |");
                    }

                    var mappings = TryGetObj(map, "Mappings") as System.Collections.IEnumerable;
                    if (mappings != null)
                    {
                        _combinedWriter.WriteLine("Mappings:");
                        foreach (var mapping in mappings)
                            _combinedWriter.WriteLine($"- {TryGetProp(mapping, "MappingTable")}");
                    }

                    _combinedWriter.WriteLine();
                    _combinedWriter.Flush();
                    count++;
                }
                catch { }
            }
            _combinedWriter.WriteLine("---");
            _combinedWriter.WriteLine();
            RecordCount("Maps", count);
        }

        // ═══════════════════════════════════════════════════════════════
        //  SECURITY (all use name-only safe pattern)
        // ═══════════════════════════════════════════════════════════════

        private void ExtractSecurityObjects()
        {
            if (ShouldExtract(ExtractionCategory.SecurityRoles))
            {
                WriteStep("Extracting Security Roles...");
                ExtractNameOnlyList("Security Roles",
                    m => MetadataProvider.SecurityRoles.ListObjects(m),
                    name => MetadataProvider.SecurityRoles.Read(name));
            }

            if (ShouldExtract(ExtractionCategory.SecurityDuties))
            {
                WriteStep("Extracting Security Duties...");
                ExtractNameOnlyList("Security Duties",
                    m => MetadataProvider.SecurityDuties.ListObjects(m),
                    name => MetadataProvider.SecurityDuties.Read(name));
            }

            if (ShouldExtract(ExtractionCategory.SecurityPrivileges))
            {
                WriteStep("Extracting Security Privileges...");
                ExtractNameOnlyList("Security Privileges",
                    m => MetadataProvider.SecurityPrivileges.ListObjects(m),
                    name => MetadataProvider.SecurityPrivileges.Read(name));
            }
        }

        // ═══════════════════════════════════════════════════════════════
        //  COMPOSITE & AGGREGATE
        // ═══════════════════════════════════════════════════════════════

        private void ExtractCompositeAndAggregate()
        {
            if (ShouldExtract(ExtractionCategory.CompositeDataEntities))
            {
                WriteStep("Extracting Composite Data Entities...");
                ExtractNameOnlyList("Composite Data Entities",
                    m => MetadataProvider.CompositeDataEntityViews.ListObjects(m),
                    name => MetadataProvider.CompositeDataEntityViews.Read(name));
            }

            if (ShouldExtract(ExtractionCategory.AggregateDataEntities))
            {
                WriteStep("Extracting Aggregate Data Entities...");
                ExtractNameOnlyList("Aggregate Data Entities",
                    m => MetadataProvider.AggregateDataEntities.ListObjects(m),
                    name => MetadataProvider.AggregateDataEntities.Read(name));
            }
        }

        /// <summary>
        /// Generic extractor for simple name + optional label categories.
        /// Uses TryGetProp so it never crashes on missing .Label.
        /// </summary>
        private void ExtractNameOnlyList(
            string categoryName,
            Func<string, IEnumerable<string>> listFunc,
            Func<string, object> readFunc)
        {
            var names = GetAllObjectNames(listFunc);

            _combinedWriter.WriteLine($"# {categoryName}");
            _combinedWriter.WriteLine();

            int count = 0;
            foreach (var name in names)
            {
                try
                {
                    var obj = readFunc(name);
                    if (obj == null) continue;

                    string objName = TryGetProp(obj, "Name");
                    string label   = TryGetProp(obj, "Label");
                    string display = !string.IsNullOrWhiteSpace(label) ? $" — {Clean(label)}" : "";
                    _combinedWriter.WriteLine($"- {objName}{display}");
                    count++;
                }
                catch { }
            }
            _combinedWriter.WriteLine();
            _combinedWriter.WriteLine("---");
            _combinedWriter.WriteLine();
            _combinedWriter.Flush();
            RecordCount(categoryName, count);
        }

    }
}

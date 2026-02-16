# CLAUDE.md

This file provides guidance to Claude Code (claude.ai/code) when working with code in this repository.

## Project Overview

D365FOMetadataExtractor is a console application that extracts metadata from Microsoft Dynamics 365 Finance and Operations (D365FO) instances. It reads metadata from the D365FO metadata provider and generates markdown files documenting tables, enums, EDTs, classes, forms, views, data entities, and other AOT objects.

## Build and Run Commands

### Build
```bash
msbuild D365FOMetadataExtractor.csproj /p:Configuration=Debug
msbuild D365FOMetadataExtractor.csproj /p:Configuration=Release
```

### Run
```bash
.\bin\Debug\D365FOMetadataExtractor.exe
.\bin\Debug\D365FOMetadataExtractor.exe "C:\CustomOutputPath"
```

The application accepts an optional command-line argument to specify the output directory (defaults to `C:\Temp\D365FO_Metadata`).

When you run the application, you'll be presented with an interactive menu to select which metadata categories to extract (Enums, Tables, Classes, etc.). You can select multiple categories by entering comma-separated numbers (e.g., `1,3,5`) or enter `0` to extract all categories.

## Architecture

### Strategy Pattern for Extraction
The codebase uses the Strategy pattern to support multiple extraction implementations:

- **ExtractionCategory Enum**: Defines 16 metadata categories (Enums, Edts, Tables, Views, DataEntities, Classes, Forms, MenuItems, Queries, Services, Maps, SecurityRoles, SecurityDuties, SecurityPrivileges, CompositeDataEntities, AggregateDataEntities)
- **IMetadataExtractor**: Interface defining the contract for all extraction strategies (`VersionName` property and `Execute()` method that accepts selected categories)
- **MetadataExtractorBase**: Abstract base class providing shared infrastructure (model iteration, console helpers, IO utilities, type resolution, category selection checking via `ShouldExtract()`)
- **DetailedMetadataExtractor**: Currently active implementation that extracts comprehensive metadata including fields, indexes, relations, methods, etc., and outputs to a **single combined markdown file** (`D365FO_Metadata.md`)

### Core Workflow (Program.cs)
1. Initialize D365FO metadata provider using `EnvironmentFactory` and `MetadataProviderFactory`
2. Discover models by scanning for Descriptor XML files in the packages directory
3. **User selects metadata categories** via interactive menu (`GetUserSelection()`)
4. Delegate to the selected `IMetadataExtractor` strategy via `Execute()` with selected categories

### Model Iteration Pattern
The proven pattern for iterating models is implemented in `MetadataExtractorBase.GetAllObjectNames()`:
- Accepts a `Func<string, IEnumerable<string>>` that lists objects for a given model
- Uses `HashSet` for deduplication (same object can appear across models)
- Try/catch per model to handle restricted/empty models gracefully
- Returns sorted list of unique object names

### Defensive Metadata Access (DetailedMetadataExtractor)
All metadata property access uses reflection-based `TryGetProp()` and `TryGetObj()` methods:
- **Why**: D365FO metadata DLL versions vary across environments; properties may not exist in all versions
- **Only .Name is accessed directly** - it's guaranteed on all metadata objects
- **All other properties** (Label, Extends, Mandatory, etc.) use `TryGetProp()` to avoid compile errors
- Returns default value (empty string or null) if property doesn't exist

### Type Resolution Helpers
Helper methods in `MetadataExtractorBase` resolve types for:
- **EDTs**: `GetEdtBaseType()` - maps `AxEdt*` subclasses to type names (String, Int, Real, Date, Enum, etc.)
- **Table Fields**: `GetFieldType()` - maps `AxTableField*` subclasses to type names
- **Map Fields**: `GetMapFieldType()` - maps `AxMapField*` subclasses to type names
- **EDT/Enum References**: `GetFieldEdtOrEnum()` - extracts ExtendedDataType or EnumType property from fields

## Output Files (DetailedMetadataExtractor)

The extractor generates **a single combined markdown file**: `D365FO_Metadata.md`

This file contains all selected metadata categories in one document, with sections separated by `---` dividers:

1. **Header** - Extraction timestamp and model count
2. **Base Enums** - Enums with values (if selected)
3. **Extended Data Types (EDTs)** - EDTs with base type, extends, length, label (if selected)
4. **Tables** - Tables with fields (EDT/Enum, type, mandatory), indexes (unique flag), relations, field groups (if selected)
5. **Views** - Views with fields (if selected)
6. **Data Entities** - Data entities with public name, collection name, fields (if selected)
7. **Classes** - Key framework classes with methods (detailed); all class names (list) (if selected)
8. **Forms** - Form names with optional labels (if selected)
9. **Menu Items** - Display/Action/Output menu items with target objects (if selected)
10. **Queries** - Queries with data sources (if selected)
11. **Services** - Services with operations (if selected)
12. **Maps** - Maps with fields and mappings (if selected)
13. **Security Roles** - Security roles (if selected)
14. **Security Duties** - Security duties (if selected)
15. **Security Privileges** - Security privileges (if selected)
16. **Composite Data Entities** - Composite data entities (if selected)
17. **Aggregate Data Entities** - Aggregate data entities (if selected)

## D365FO Dependencies

The project references Microsoft Dynamics AX metadata DLLs from the AOSService packages directory:
- `Microsoft.Dynamics.ApplicationPlatform.Environment.dll`
- `Microsoft.Dynamics.AX.Metadata.dll`
- `Microsoft.Dynamics.AX.Metadata.Core.dll`
- `Microsoft.Dynamics.AX.Metadata.Storage.dll`

**HintPath**: `..\..\..\..\..\AOSService\PackagesLocalDirectory\bin\`

This application must run on a machine with D365FO installed, or the DLL references must be updated to valid paths.

## Key Extension Points

### Adding a New Extraction Strategy
1. Create a class implementing `IMetadataExtractor`
2. Inherit from `MetadataExtractorBase` for shared infrastructure
3. Implement `VersionName` property and `RunExtraction()` method
4. Use `GetAllObjectNames()` for model iteration
5. Use `TryGetProp()`/`TryGetObj()` for safe property access
6. Use `RecordCount()` to track extracted objects for summary
7. Use `ShouldExtract(ExtractionCategory)` to check if a category was selected by the user
8. Update `Program.cs` to instantiate the new strategy

### Adding a New Object Type
1. Add new value to `ExtractionCategory` enum in `IMetadataExtractor.cs`
2. Update selection menu in `Program.GetUserSelection()` to include the new category
3. Add extraction method in `DetailedMetadataExtractor` (e.g., `ExtractReports()`)
4. Call it from `RunExtraction()` with conditional check: `if (ShouldExtract(ExtractionCategory.Reports)) ExtractReports();`
5. Write to `_combinedWriter` instead of creating a new file
6. Use pattern: `GetAllObjectNames(m => MetadataProvider.{Category}.ListObjects(m))`
7. Use `TryGetProp()` for all property access except `.Name`
8. Add `---` separator at end of method
9. Call `RecordCount()` with category name and count

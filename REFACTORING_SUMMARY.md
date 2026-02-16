# Refactoring Summary

## Changes Made

### 1. Added User Selection Menu
- Modified `Program.cs` to include an interactive selection menu
- Users can now choose which metadata categories to extract (1-16 or 0 for all)
- Multiple selections supported via comma-separated input
- Selected categories are displayed before extraction begins

### 2. Updated Interface (IMetadataExtractor.cs)
- Added `ExtractionCategory` enum with 16 categories:
  - Enums, Edts, Tables, Views, DataEntities, Classes, Forms
  - MenuItems, Queries, Services, Maps
  - SecurityRoles, SecurityDuties, SecurityPrivileges
  - CompositeDataEntities, AggregateDataEntities
- Modified `Execute()` method signature to accept `HashSet<ExtractionCategory> selectedCategories`

### 3. Updated Base Class (MetadataExtractorBase.cs)
- Added `SelectedCategories` property to track user selections
- Added `ShouldExtract(ExtractionCategory category)` helper method
- Updated `Execute()` method to accept and store selected categories

### 4. Refactored DetailedMetadataExtractor.cs
**Major Changes:**
- **Single Output File**: All metadata now writes to `D365FO_Metadata.md` instead of 17 separate files
- **Combined Writer**: Introduced `_combinedWriter` field used across all extraction methods
- **Conditional Extraction**: Each category only extracts if user selected it via `ShouldExtract()`
- **Section Separators**: Added `---` markdown separators between sections

**Updated Methods:**
- `RunExtraction()`: Opens single file writer, conditionally calls extraction methods
- `ExtractEnums()`: Writes to combined file
- `ExtractEdts()`: Writes to combined file
- `ExtractTables()`: Writes to combined file (with field names, indexes, relations)
- `ExtractViews()`: Writes to combined file
- `ExtractDataEntities()`: Writes to combined file
- `ExtractClasses()`: Writes to combined file
- `ExtractForms()`: Writes to combined file
- `ExtractMenuItems()`: Writes to combined file (all three types: Display, Action, Output)
- `ExtractQueries()`: Writes to combined file
- `ExtractServices()`: Writes to combined file
- `ExtractMaps()`: Writes to combined file
- `ExtractSecurityObjects()`: Conditionally extracts Roles, Duties, Privileges
- `ExtractCompositeAndAggregate()`: Conditionally extracts Composite and Aggregate entities
- `ExtractNameOnlyList()`: Updated signature (removed fileName parameter), writes to combined file

**Removed Methods:**
- `GenerateCombinedSkillFile()`: No longer needed (single file already contains everything)
- `GenerateIndex()`: No longer needed (header section includes metadata)

## Output Changes

### Before:
- 17+ separate markdown files (01_BaseEnums.md, 02_ExtendedDataTypes.md, etc.)
- All categories extracted every time
- No user interaction

### After:
- **Single file**: `D365FO_Metadata.md`
- **User-selected categories only**
- **Interactive menu** at startup
- **Sections separated** by `---` dividers

## Usage Example

```
> D365FOMetadataExtractor.exe "C:\Output"

╔══════════════════════════════════════════════════════╗
║           Select Metadata to Extract                ║
╚══════════════════════════════════════════════════════╝

Available categories:
  1.  Enums
  2.  Extended Data Types (EDTs)
  3.  Tables
  ...
  0.  All categories

Enter selections (comma-separated, e.g., 1,3,5 or 0 for all): 1,2,3

Selected 3 category(ies):
  - Enums
  - Edts
  - Tables

> Running: V2 — Detailed (Fields, Indexes, Relations) - Combined Output
> Extracting Base Enums (with values)...
> Extracting Extended Data Types...
> Extracting Tables (fields, indexes, relations — may take several minutes)...
...
```

## Benefits

1. **Faster Extraction**: Only extract what you need
2. **Easier Navigation**: Single file easier to search and reference
3. **Flexible**: Choose different combinations for different needs
4. **User-Friendly**: Interactive menu instead of code changes
5. **Clean Output**: One file instead of 17+ scattered files

## Backward Compatibility

To extract everything like before, simply enter `0` when prompted, which selects all categories.

## Build Notes

The refactored code maintains all existing functionality and uses the same defensive programming patterns:
- Reflection-based `TryGetProp()` for cross-version compatibility
- Try-catch blocks per object to handle errors gracefully
- Model iteration pattern unchanged
- All type resolution helpers preserved

# D365FO Metadata Extractor

A console application that extracts metadata from Microsoft Dynamics 365 Finance and Operations (D365FO) instances and generates comprehensive markdown documentation.

## Overview

D365FOMetadataExtractor reads metadata from the D365FO metadata provider and generates markdown files documenting:
- Tables, Views, and Data Entities
- Enums and Extended Data Types (EDTs)
- Classes, Forms, and Menu Items
- Queries, Services, and Maps
- Security Roles, Duties, and Privileges
- Composite and Aggregate Data Entities

## Features

- **Interactive Category Selection**: Choose which metadata categories to extract via menu
- **Single Combined Output**: All metadata exported to one markdown file (`D365FO_Metadata.md`)
- **Comprehensive Details**: Extracts fields, indexes, relations, methods, and more
- **Multi-Model Support**: Processes all models in your D365FO instance
- **Defensive Programming**: Uses reflection-based property access for version compatibility

## Prerequisites

- Microsoft Dynamics 365 Finance and Operations installed
- .NET Framework 4.6.1 or higher
- Access to D365FO metadata DLLs

## Building

```bash
msbuild D365FOMetadataExtractor.csproj /p:Configuration=Release
```

## Usage

```bash
.\bin\Debug\D365FOMetadataExtractor.exe
.\bin\Debug\D365FOMetadataExtractor.exe "C:\CustomOutputPath"
```

The application accepts an optional command-line argument to specify the output directory (defaults to `C:\Temp\D365FO_Metadata`).

When you run the application, you'll see an interactive menu to select metadata categories:
- Enter comma-separated numbers (e.g., `1,3,5`) to select specific categories
- Enter `0` to extract all categories

## Architecture

The codebase uses the **Strategy Pattern** for metadata extraction:
- `IMetadataExtractor`: Interface for extraction strategies
- `MetadataExtractorBase`: Abstract base with shared infrastructure
- `DetailedMetadataExtractor`: Current implementation with comprehensive extraction

See [CLAUDE.md](CLAUDE.md) for detailed architecture and extension guidance.

## Output

Generates `D365FO_Metadata.md` containing all selected metadata categories with sections for enums, tables, classes, forms, and more.

## License

[Specify your license here]

## Contributing

[Add contribution guidelines if needed]

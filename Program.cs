// ============================================================================
// Program.cs — Thin Orchestrator
// ============================================================================
// Responsibilities:
//   1. Initialize the D365FO metadata provider (proven pattern)
//   2. Discover models (Descriptor XML scan)
//   3. Delegate to the selected IMetadataExtractor strategy
//
// Switch versions by changing ONE LINE in Main().
// ============================================================================

using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.IO;
using System.Linq;
using Microsoft.Dynamics.ApplicationPlatform.Environment;
using Microsoft.Dynamics.AX.Metadata.Providers;
using Microsoft.Dynamics.AX.Metadata.Storage;
using Microsoft.Dynamics.AX.Metadata.Storage.Runtime;

namespace D365FOMetadataExtractor
{
    class Program
    {
        static void Main (string[] args)
        {
            IMetadataExtractor extractor = new DetailedMetadataExtractor();
            string outputDirectory = @"C:\Temp\D365FO_Metadata";

            // ── Override from command line if provided ──
            if (args.Length >= 1) outputDirectory = args[0];

            // ── Banner ──
            Console.ForegroundColor = ConsoleColor.Cyan;
            Console.WriteLine("╔══════════════════════════════════════════════════════╗");
            Console.WriteLine("║        D365FO Metadata Extractor                    ║");
            Console.WriteLine("╚══════════════════════════════════════════════════════╝");
            Console.ResetColor();
            Console.WriteLine($"  Strategy : {extractor.VersionName}");
            Console.WriteLine($"  Output   : {outputDirectory}");
            Console.WriteLine();

            var totalTimer = Stopwatch.StartNew();

            try
            {
                // ══════════════════════════════════════════════════════
                //  STEP 1: Initialize provider (proven working pattern)
                // ══════════════════════════════════════════════════════
                Console.ForegroundColor = ConsoleColor.Yellow;
                Console.WriteLine("> Initializing D365FO Metadata Provider...");
                Console.ResetColor();

                var environment = EnvironmentFactory.GetApplicationEnvironment();
                var packagesDir = environment.Aos.PackageDirectory;
                Console.WriteLine($"  Packages: {packagesDir}");

                var runtimeConfig = new RuntimeProviderConfiguration(packagesDir);
                IMetadataProvider metadataProvider = new MetadataProviderFactory()
                    .CreateRuntimeProviderWithExtensions(runtimeConfig);

                Console.ForegroundColor = ConsoleColor.Green;
                Console.WriteLine("  Provider initialized.");
                Console.ResetColor();

                // ══════════════════════════════════════════════════════
                //  STEP 2: Discover models (Descriptor XML scan)
                // ══════════════════════════════════════════════════════
                Console.ForegroundColor = ConsoleColor.Yellow;
                Console.WriteLine("\n> Discovering models...");
                Console.ResetColor();

                // The metadata provider API uses PACKAGE NAMES (e.g., "ApplicationSuite"),
                // not descriptor file names (e.g., "Foundation.xml").
                // We need to get the package directory names, not the descriptor XML filenames.
                List<string> modelNames = Directory.GetDirectories(packagesDir)
                    .Where(packageDir => Directory.Exists(Path.Combine(packageDir, "Descriptor")))
                    .Select(packageDir => Path.GetFileName(packageDir))
                    .Distinct()
                    .OrderBy(n => n)
                    .ToList();

                Console.WriteLine($"  Found {modelNames.Count} models:");
                foreach (var modelName in modelNames.Take(20))
                {
                    Console.WriteLine($"    - {modelName}");
                }
                if (modelNames.Count > 20)
                {
                    Console.WriteLine($"    ... and {modelNames.Count - 20} more");
                }

                // Validation check for ApplicationSuite package
                if (modelNames.Contains("ApplicationSuite"))
                {
                    Console.ForegroundColor = ConsoleColor.Green;
                    Console.WriteLine("\n  ✓ ApplicationSuite package found in list");
                    Console.ResetColor();
                }
                else
                {
                    Console.ForegroundColor = ConsoleColor.Red;
                    Console.WriteLine("\n  ✗ ApplicationSuite package NOT found - standard tables will be missing!");
                    Console.ResetColor();
                }

                // ══════════════════════════════════════════════════════
                //  STEP 3: User selection of metadata categories
                // ══════════════════════════════════════════════════════
                HashSet<ExtractionCategory> selectedCategories = GetUserSelection();

                if (selectedCategories.Count == 0)
                {
                    Console.ForegroundColor = ConsoleColor.Red;
                    Console.WriteLine("\nNo categories selected. Exiting...");
                    Console.ResetColor();
                    Console.WriteLine("\nPress Enter to exit...");
                    Console.ReadLine();
                    return;
                }

                // ══════════════════════════════════════════════════════
                //  STEP 4: Delegate to selected strategy
                // ══════════════════════════════════════════════════════
                Console.ForegroundColor = ConsoleColor.Yellow;
                Console.WriteLine($"\n> Running: {extractor.VersionName}");
                Console.ResetColor();

                extractor.Execute(metadataProvider, modelNames, outputDirectory, selectedCategories);
            }
            catch (Exception ex)
            {
                Console.ForegroundColor = ConsoleColor.Red;
                Console.WriteLine($"\nFATAL ERROR: {ex.Message}");
                Console.WriteLine(ex.StackTrace);
                Console.ResetColor();
            }

            totalTimer.Stop();
            Console.WriteLine();
            Console.ForegroundColor = ConsoleColor.Green;
            Console.WriteLine($"══ Done in {totalTimer.Elapsed.TotalMinutes:F1} minutes ══");
            Console.WriteLine($"   Output: {Path.GetFullPath(outputDirectory)}");
            Console.ResetColor();
            Console.WriteLine("\nPress Enter to exit...");
            Console.ReadLine();
        }

        static HashSet<ExtractionCategory> GetUserSelection()
        {
            Console.ForegroundColor = ConsoleColor.Cyan;
            Console.WriteLine("\n╔══════════════════════════════════════════════════════╗");
            Console.WriteLine("║           Select Metadata to Extract                ║");
            Console.WriteLine("╚══════════════════════════════════════════════════════╝");
            Console.ResetColor();
            Console.WriteLine();
            Console.WriteLine("Available categories:");
            Console.WriteLine("  1.  Enums");
            Console.WriteLine("  2.  Extended Data Types (EDTs)");
            Console.WriteLine("  3.  Tables");
            Console.WriteLine("  4.  Views");
            Console.WriteLine("  5.  Data Entities");
            Console.WriteLine("  6.  Classes");
            Console.WriteLine("  7.  Forms");
            Console.WriteLine("  8.  Menu Items");
            Console.WriteLine("  9.  Queries");
            Console.WriteLine("  10. Services");
            Console.WriteLine("  11. Maps");
            Console.WriteLine("  12. Security Roles");
            Console.WriteLine("  13. Security Duties");
            Console.WriteLine("  14. Security Privileges");
            Console.WriteLine("  15. Composite Data Entities");
            Console.WriteLine("  16. Aggregate Data Entities");
            Console.WriteLine("  0.  All categories");
            Console.WriteLine();
            Console.Write("Enter selections (comma-separated, e.g., 1,3,5 or 0 for all): ");

            string input = Console.ReadLine();
            var selectedCategories = new HashSet<ExtractionCategory>();

            if (string.IsNullOrWhiteSpace(input))
            {
                return selectedCategories;
            }

            var selections = input.Split(new[] { ',', ' ' }, StringSplitOptions.RemoveEmptyEntries);

            foreach (var selection in selections)
            {
                if (int.TryParse(selection.Trim(), out int choice))
                {
                    if (choice == 0)
                    {
                        // Select all categories
                        foreach (ExtractionCategory category in Enum.GetValues(typeof(ExtractionCategory)))
                        {
                            selectedCategories.Add(category);
                        }
                        break;
                    }
                    else if (Enum.IsDefined(typeof(ExtractionCategory), choice))
                    {
                        selectedCategories.Add((ExtractionCategory)choice);
                    }
                    else
                    {
                        Console.ForegroundColor = ConsoleColor.Yellow;
                        Console.WriteLine($"  Warning: Invalid selection '{choice}' - skipped");
                        Console.ResetColor();
                    }
                }
            }

            Console.WriteLine();
            Console.ForegroundColor = ConsoleColor.Green;
            Console.WriteLine($"Selected {selectedCategories.Count} category(ies):");
            foreach (var category in selectedCategories.OrderBy(c => (int)c))
            {
                Console.WriteLine($"  - {category}");
            }
            Console.ResetColor();

            return selectedCategories;
        }
    }
}
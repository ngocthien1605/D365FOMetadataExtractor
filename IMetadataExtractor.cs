// ============================================================================
// IMetadataExtractor.cs — Strategy Interface
// ============================================================================
// Both V1 (basic) and V2 (detailed) extractors implement this contract.
// Program.Main() picks which strategy to run.
// ============================================================================

using Microsoft.Dynamics.AX.Metadata.Providers;
using System.Collections.Generic;

namespace D365FOMetadataExtractor
{
    /// <summary>
    /// Metadata extraction categories.
    /// </summary>
    public enum ExtractionCategory
    {
        Enums = 1,
        Edts = 2,
        Tables = 3,
        Views = 4,
        DataEntities = 5,
        Classes = 6,
        Forms = 7,
        MenuItems = 8,
        Queries = 9,
        Services = 10,
        Maps = 11,
        SecurityRoles = 12,
        SecurityDuties = 13,
        SecurityPrivileges = 14,
        CompositeDataEntities = 15,
        AggregateDataEntities = 16
    }

    /// <summary>
    /// Contract for metadata extraction strategies.
    /// </summary>
    public interface IMetadataExtractor
    {
        /// <summary>
        /// Human-readable name shown in console output.
        /// </summary>
        string VersionName { get; }

        /// <summary>
        /// Runs the full extraction pipeline.
        /// </summary>
        /// <param name="metadataProvider">Initialized D365FO metadata provider.</param>
        /// <param name="modelNames">Discovered model names from Descriptor scan.</param>
        /// <param name="outputDirectory">Root folder for all output files.</param>
        /// <param name="selectedCategories">Categories to extract (null or empty = all).</param>
        void Execute(IMetadataProvider metadataProvider, List<string> modelNames, string outputDirectory, HashSet<ExtractionCategory> selectedCategories);
    }
}

<!-- Use this file to provide workspace-specific custom instructions to Copilot. For more details, visit https://code.visualstudio.com/docs/copilot/copilot-customization#_use-a-githubcopilotinstructionsmd-file -->

# PowerShell Application Comparison Toolkit - Copilot Instructions

This workspace contains PowerShell scripts for comparing installed applications between Windows computers.

## Project Context
- **Purpose**: Compare software installations across multiple Windows computers
- **Target Users**: IT administrators, system administrators, technical support teams
- **Platform**: Windows PowerShell 5.1+ and PowerShell 7+

## Code Standards
- Follow PowerShell best practices and coding standards
- Use approved PowerShell verbs (Get-, Set-, New-, Remove-, etc.)
- Include comprehensive error handling with try-catch blocks
- Implement proper parameter validation and type constraints
- Use consistent formatting and indentation (4 spaces)
- Include detailed comment-based help for functions

## Script Architecture
- **Export-InstalledApps.ps1**: Collects application data from multiple sources (Registry, WMI, Package Managers)
- **Compare-InstalledApps.ps1**: Analyzes differences between two application datasets
- **Start-AppComparison.ps1**: Provides guided workflow for complete comparison process

## Key Features to Maintain
- Multi-source application detection (Registry, WMI, AppX, Package Managers)
- Comprehensive error handling and logging
- HTML report generation with visual formatting
- JSON export capabilities for programmatic use
- Similar application grouping functionality
- Version difference detection and highlighting

## PowerShell Specific Guidelines
- Use `[CmdletBinding()]` for advanced function parameters
- Implement proper parameter sets and validation
- Use `Write-Verbose`, `Write-Warning`, and `Write-Error` appropriately
- Follow PowerShell naming conventions (Verb-Noun pattern)
- Include parameter help text and examples
- Use `PSCustomObject` for structured data output

## Performance Considerations
- Registry-based detection should be primary method (fastest)
- WMI queries can be slow, handle timeouts gracefully
- Implement deduplication logic for overlapping data sources
- Consider memory usage when processing large application lists

## Security Notes
- Scripts should only read system information, never modify
- Handle access denied errors gracefully
- No sensitive data collection (credentials, personal files, etc.)
- Consider execution policy requirements in documentation

## HTML Report Standards
- Responsive design that works on different screen sizes
- Clear color coding for different comparison categories
- Accessible markup with proper semantic HTML
- Professional styling appropriate for business reports
- Include summary statistics and computer identification

## Data Format Standards
- Use JSON for data interchange between scripts
- Include metadata (computer name, export date, options used)
- Consistent property naming across all data structures
- Handle null/empty values appropriately
- Maintain backward compatibility for data format changes

When suggesting improvements or modifications, prioritize:
1. Reliability and error handling
2. Performance optimization
3. User experience and clarity
4. Maintainability and code quality
5. Cross-version PowerShell compatibility

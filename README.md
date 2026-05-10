# Package Creator Tool

A comprehensive GUI-based tool for creating standardized software package structures and deployment scripts for enterprise software distribution via Microsoft Endpoint Configuration Manager (SCCM/MECM) and Microsoft Intune.

## Description

The Package Creator Tool simplifies the process of preparing software packages for deployment in enterprise environments. It provides a user-friendly Windows Forms interface that guides users through the package creation process, automatically generates necessary scripts, and organizes files in a standardized structure.

The tool supports both MSI and EXE installers, with intelligent metadata extraction for MSI files, and generates deployment scripts compatible with SCCM/MECM and Intune packaging requirements.

## Features

### GUI Interface
- Intuitive Windows Forms interface for package configuration
- Input validation for required fields
- Browse functionality for selecting installer files

### Package Creation
- **Standardized Folder Structure**: Creates organized package directories following naming conventions
- **Metadata Extraction**: Automatically extracts MSI properties (product code, version, manufacturer, etc.)
- **Detection Methods**: Configurable detection rules for MSI (product code) and EXE (file-based detection)
- **Zone Support**: Support for zone-specific packages (MAIN, ZN01-ZN15)

### Generated Scripts and Resources
- **PSADT Integration**: Uses PowerShell App Deployment Toolkit templates
- **Test Scripts**: Generates install/uninstall test scripts using PsExec
- **SCCM Deployment**: Creates full SCCM application deployment scripts
- **Intune Conversion**: Generates scripts for converting packages to Intune .intunewin format
- **Documentation Templates**: Includes documentation templates for package maintenance

### Supported Installers
- MSI files with automatic property extraction
- EXE files with manual detection configuration
- Option to include entire installer source folders

## Prerequisites

- Windows 10/11 or Windows Server
- PowerShell 5.1 or higher
- .NET Framework (for Windows Forms)
- Access to SCCM/MECM console (for deployment)
- Microsoft IntuneWinAppUtil.exe (included in templates)

## Installation

1. Clone or download the repository to your local machine
2. Ensure the folder structure matches the expected layout
3. Run the script with appropriate permissions

## Usage

### Basic Usage

```powershell
.\Create_FSv2.ps1
```

Or specify a custom root path:

```powershell
.\Create_FSv2.ps1 'C:\Packages'
```

### GUI Workflow

1. **Launch the Tool**: Run the script to open the Package Creator GUI
2. **Configure Package Details**:
   - Enter manufacturer, product name, version
   - Select architecture (x86/x64/All)
   - Specify language and revision
   - Choose zone code if applicable
3. **Select Installer**:
   - Browse for MSI or EXE file
   - For MSI: metadata is auto-extracted
   - For EXE: manually configure detection rules
4. **Set Detection Methods**:
   - MSI: Uses product code
   - EXE: Specify file path, name, and version
5. **Create Package**: Click "Create package" to generate the complete package structure

### Generated Structure

The tool creates packages in the following structure:

```
RootPath\
└── Vendor\
    └── Product\
        └── Version_Arch_Lang_Rev_Zone\
            ├── Documentation\
            │   └── PackageName - Documentation.docx
            ├── Resources\
            │   ├── test_Install.ps1
            │   ├── test_Uninstall.ps1
            │   ├── Deploy_SCCM.ps1
            │   └── Convert_Intunewin.ps1
            └── Package\
                ├── Deploy-Application.ps1
                ├── Deploy-Application.exe.config
                ├── Files\
                │   └── [installer files]
                └── AppDeployToolkit\
                    └── [PSADT files]
```

## Configuration

The tool stores user preferences in Windows Registry under `HKCU:\SOFTWARE\MyCompany\PackageCreator\`:
- TargetFolder: Default root path
- MSIVer: Tool version
- SCCM_SiteCode: Default SCCM site code
- SCCM_Repository: Default SCCM source path

## Templates

The tool uses templates located in `Templates\`:
- `Package\`: PSADT base files
- `Deploy_template.ps1`: SCCM deployment script template
- `Convert_Intunewin_template.ps1`: Intune conversion script template
- `PackageName - Documentation.docx`: Documentation template

## Troubleshooting

### Common Issues

1. **Permission Errors**: Ensure you have write access to the target folder
2. **Missing Templates**: Verify all template files are present in the Templates directory
3. **Registry Access**: The tool requires access to HKCU registry for settings

### Validation Checks

The tool validates:
- Architecture selection (at least one must be selected)
- Required fields (Manufacturer, ProductName, Version, Language, Revision)
- Folder creation permissions

## Contributing

1. Fork the repository
2. Create a feature branch
3. Make your changes
4. Test thoroughly
5. Submit a pull request

## Version History

- **v4.3** (2026-05-10): Updated documentation and minor GUI improvements
- **v4.2** (2024-05-20): Previous version with core functionality

## License

[Specify license if applicable]

## Support

For support or questions, please [contact method or create an issue].

## Author

Viktor Kravtsov</content>

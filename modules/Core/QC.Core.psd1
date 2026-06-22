@{
    RootModule        = ''
    ModuleVersion     = '0.1.0'
    GUID              = 'be3d1b45-ef02-4f08-a328-2c8b8f5f9b6f'
    Author            = 'TYPSA'
    CompanyName       = 'TYPSA'
    Copyright         = '(c) TYPSA. Prototype manifest; not used by production entrypoints.'
    Description       = 'Prototype manifest bundling QC core modules (Results, Runtime, Paths, Config, Logging, Hashing, Telemetry, WatcherOrchestration).'

    PowerShellVersion = '5.1'

    NestedModules     = @(
        'Core.Results.psm1'
        'Core.Runtime.psm1'
        'Core.Paths.psm1'
        'Core.Config.psm1'
        'Core.Logging.psm1'
        'Core.Hashing.psm1'
        'Core.Telemetry.psm1'
        'QC.WatcherOrchestration.psm1'
    )

    # Prototype: wildcard exports preserve nested-module surface area. Narrow in a later phase.
    FunctionsToExport = '*'
    CmdletsToExport     = @()
    VariablesToExport   = '*'
    AliasesToExport     = @()

    PrivateData         = @{
        PSData = @{
            Tags         = @('ProjectWise', 'QC', 'Prototype', 'Core')
            ProjectUri   = ''
            ReleaseNotes = 'Phase 4G prototype only. Production entrypoints still import folder .psm1 paths directly.'
        }
    }
}

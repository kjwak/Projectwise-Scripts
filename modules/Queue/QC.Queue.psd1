@{
    RootModule        = ''
    ModuleVersion     = '0.1.0'
    GUID              = '22fef31f-8f0d-428d-8067-45bb884641ca'
    Author            = 'TYPSA'
    CompanyName       = 'TYPSA'
    Copyright         = '(c) TYPSA. Prototype manifest; not used by production entrypoints.'
    Description       = 'Prototype manifest bundling QC queue modules (Filters, Triggers, JobFactory, Queue.Json, Worker).'

    PowerShellVersion = '5.1'

    NestedModules     = @(
        'QC.Filters.psm1'
        'QC.Triggers.psm1'
        'QC.JobFactory.psm1'
        'QC.Queue.Json.psm1'
        'QC.Worker.psm1'
    )

    # Prototype: wildcard exports preserve nested-module surface area. Narrow in a later phase.
    FunctionsToExport = '*'
    CmdletsToExport     = @()
    VariablesToExport   = '*'
    AliasesToExport     = @()

    PrivateData         = @{
        PSData = @{
            Tags         = @('ProjectWise', 'QC', 'Prototype', 'Queue')
            ProjectUri   = ''
            ReleaseNotes = 'Phase 4G prototype only. Import QC.Core.psd1 first in tests; production still uses folder .psm1 paths.'
        }
    }
}

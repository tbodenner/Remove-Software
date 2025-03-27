# I have not tested this yet

function Add-DllFix {
    # Add the new DLLs to the Global Assembly Cache
    Add-Type -AssemblyName "System.EnterpriseServices"
    $publish = [System.EnterpriseServices.Internal.Publish]::new()

    $dlls = @(
        'System.Memory.dll',
        'System.Numerics.Vectors.dll',
        'System.Runtime.CompilerServices.Unsafe.dll',
        'System.Security.Principal.Windows.dll'
    )

    foreach ($dll in $dlls) {
        $dllPath = "$env:SystemRoot\System32\WindowsPowerShell\v1.0\$dll"
        $publish.GacInstall($dllPath)
    }    

    # Create a file so we can easily track that this computer was fixed (in case we need to revert)
    New-Item -Path "$env:SystemRoot\System32\WindowsPowerShell\v1.0\" -Name DllFix.txt -ItemType File -Value "$dlls added to the Global Assembly Cache"
    #Restart-Computer
}

function Remove-DllFix {
    if (Test-Path "$env:SystemRoot\System32\WindowsPowerShell\v1.0\DllFix.txt") {

        Add-Type -AssemblyName "System.EnterpriseServices"
        $publish = [System.EnterpriseServices.Internal.Publish]::new()

        $dlls = @(
            'System.Memory.dll',
            'System.Numerics.Vectors.dll',
            'System.Runtime.CompilerServices.Unsafe.dll',
            'System.Security.Principal.Windows.dll'
        )

        foreach ($dll in $dlls) {
            $dllPath = "$env:SystemRoot\System32\WindowsPowerShell\v1.0\$dll"
            $publish.GacRemove($dllPath)
        } 
    }

    Remove-Item -Path "$env:SystemRoot\System32\WindowsPowerShell\v1.0\DllFix.txt" -Force
    #Restart-Computer
}
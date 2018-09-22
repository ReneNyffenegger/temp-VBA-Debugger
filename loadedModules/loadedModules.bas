option explicit

declare function loadedModules lib "c:\temp\temp-VBA-Debugger\loadedModules\loadedModules.dll" () as string()

sub main() ' {

    dim modules() as string
    dim i         as integer

    modules = loadedModules

    for i = 0 to uBound(loadedModules)
        debug.print modules(i)
    next i


end sub ' }

option explicit

declare function loadedModules lib "c:\temp\temp-VBA-Debugger\loadedModules\loadedModules.dll" () as variant()

sub main() ' {

    dim modules() as variant
    dim i         as integer

    modules = loadedModules

    for i = 0 to uBound(loadedModules)
        debug.print modules(i)
    next i


end sub ' }

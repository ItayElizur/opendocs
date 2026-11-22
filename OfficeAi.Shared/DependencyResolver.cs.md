# DependencyResolver.cs

The add-ins ship System.Text.Json 8 and its dependencies next to the add-in DLL. Their assembly versions do not
all agree: `System.Memory` and `System.Threading.Tasks.Extensions` ask for
`System.Runtime.CompilerServices.Unsafe` 4.0.4.1, while the shipped file (the one `System.Text.Json` asks for) is
6.0.0.0. The build writes a binding redirect for it into `<AddIn>.dll.config`.

On an offline machine installed from the MSI, every app showed
`WebView2 init failed: The type initializer for 'System.Text.Json.JsonSerializer' threw an exception` at startup.
That machine had the right files: loading them from the install folder with the redirect applied serialized
fine. Loading them without the redirect gave the same `TypeInitializationException`, caused by
`FileNotFoundException` for `Unsafe` 4.0.4.1. So Office loaded the add-in there without reading its `.dll.config`.
Why it did that was not found.

`AssemblyResolve` runs only after .NET's own binding has failed, so on a machine where the redirect works this
code never runs. When it does run, it returns the assembly of that name already loaded, or else the add-in
folder's copy, whatever version was asked for. An assembly that is not in the add-in folder (for example a
`.resources` satellite) still fails as before.

It is registered from a static constructor on each app's `ThisAddIn`, which runs before VSTO creates the add-in
and its ribbon, so before any code that touches System.Text.Json. `AppDomain.CurrentDomain.BaseDirectory` is the
add-in's own folder (VSTO sets it; `ModePolicy` and `RibbonBase` rely on the same thing).

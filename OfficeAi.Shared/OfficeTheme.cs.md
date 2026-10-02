## OfficeTheme

```
/// Resolves Office's own "Office Theme" setting (File &gt; Account) to a
/// plain "dark"/"light" verdict, read once per pane at creation time -
/// there is no supported object-model property for this, only an
/// undocumented registry value, so every failure mode here degrades to
/// "light" rather than throwing (a theme-detection bug must never break
/// pane creation).
```

## OfficeTheme.ReadEffectiveTheme

```
/// Test seam: matches Microsoft.Win32.Registry.GetValue's exact
/// signature (keyName, valueName, defaultValue) -&gt; object, so the
/// real read is just `registryGetValue ?? Registry.GetValue`. Kept
/// as a plain delegate parameter rather than an interface, matching
/// this repo's existing convention (e.g. WebViewBridgeHost's
/// constructor).
```

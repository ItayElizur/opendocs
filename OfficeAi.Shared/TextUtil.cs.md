## TextUtil - seam rules

```
/// Pure string helpers shared by the Word/Excel/PowerPoint tool layers.
/// Extracted here (Phase 0) so they are unit-testable at all - the
/// *Tools.cs files live in VSTO projects whose private members no test
/// project can reach.
///
/// SEAM RULES for this assembly (OfficeAi.Shared) - read before adding
/// anything new here:
///
/// 1. Anything free of COM/Office-interop types goes here, and gets a
///    test. Anything touching Word.*/Excel.*/PowerPoint.* stays in its
///    app's own *Tools.cs - this assembly has no interop seam for actual
///    COM calls yet (that is a future phase, not this one). Without this
///    rule written down, the next helper gets added to a *Tools.cs out of
///    habit and this seam quietly stops growing.
///
/// 2. NEVER expose an Office interop type as a generic type argument from
///    this assembly - e.g. Dictionary&lt;string, MsoAutoShapeType&gt; as a
///    public member. This assembly embeds the Office PIA
///    (EmbedInteropTypes=true, same as every app project), and an
///    embedded interop type used as a generic type argument cannot cross
///    an assembly boundary - confirmed via a spike while building
///    ShapeTypes.cs (Phase 0 Task 4): CS1769, "cannot be used across
///    assembly boundaries because it has a generic type argument that is
///    an embedded interop type." The type is fine used BARE (a plain
///    parameter or return value, not inside a generic), just never as a
///    generic argument.
///    Carry it as int/string instead, and cast at the app-side call site
///    - see ColorUtil.HexToOle and ShapeTypes.ByName, both int-valued for
///    exactly this reason. This costs a build cycle to rediscover if you
///    don't already know it; don't make the next person hit it head-on.
```

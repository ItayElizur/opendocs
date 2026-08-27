using System;
using System.Collections.Generic;
using MsoAutoShapeType = Microsoft.Office.Core.MsoAutoShapeType;

namespace OfficeAi.Shared
{
    /// <summary>
    /// Shape-name to MsoAutoShapeType, as int. It is int and not
    /// MsoAutoShapeType deliberately: this assembly embeds the Office PIA,
    /// and an embedded interop type used as a GENERIC TYPE ARGUMENT cannot
    /// cross an assembly boundary (verified via a spike: CS1769,
    /// "cannot be used across assembly boundaries because it has a generic
    /// type argument that is an embedded interop type"). Callers cast:
    ///     (Microsoft.Office.Core.MsoAutoShapeType)ShapeTypes.ByName["rect"]
    /// Same split as ColorUtil uses for color - this assembly never exposes
    /// an Office interop type, only plain int/string; the app casts at the
    /// call site.
    ///
    /// Each value below is (int)MsoAutoShapeType.xxx, cast at initialization
    /// time - the enum is used bare here (Probe B in the spike), never as a
    /// Dictionary's generic type argument, so this compiles. Deliberately NOT
    /// hand-typed magic numbers - the enum member names are the source of
    /// truth and the compiler resolves the actual values.
    ///
    /// Union of Excel's and PowerPoint's previously-separate maps - PowerPoint
    /// additionally aliased "rectangle"/"oval" to "rect"/"ellipse" (kept so
    /// existing calls keep working); every other entry was already identical
    /// between the two (verified by diff).
    ///
    /// Note: the brief's source table also lists "plus"/"mathPlus"
    /// (MsoAutoShapeType.msoShapePlus/msoShapeMathPlus), but those enum
    /// members do not exist in this project's referenced Microsoft.Office.Core
    /// PIA (confirmed via CS0117 compile failure) - omitted; requests for
    /// either shape type fall back to msoShapeRectangle per the table's
    /// existing fallback behavior.
    ///
    /// "textbox" is handled separately by each app (not an AutoShape) and is
    /// not in this map.
    /// </summary>
    public static class ShapeTypes
    {
        public static readonly Dictionary<string, int> ByName =
            new Dictionary<string, int>(StringComparer.OrdinalIgnoreCase)
        {
            ["rect"] = (int)MsoAutoShapeType.msoShapeRectangle,
            ["rectangle"] = (int)MsoAutoShapeType.msoShapeRectangle,      // alias (PowerPoint)
            ["roundRect"] = (int)MsoAutoShapeType.msoShapeRoundedRectangle,
            ["ellipse"] = (int)MsoAutoShapeType.msoShapeOval,
            ["oval"] = (int)MsoAutoShapeType.msoShapeOval,                // alias (PowerPoint)
            ["triangle"] = (int)MsoAutoShapeType.msoShapeIsoscelesTriangle,
            ["rtTriangle"] = (int)MsoAutoShapeType.msoShapeRightTriangle,
            ["parallelogram"] = (int)MsoAutoShapeType.msoShapeParallelogram,
            ["trapezoid"] = (int)MsoAutoShapeType.msoShapeTrapezoid,
            ["diamond"] = (int)MsoAutoShapeType.msoShapeDiamond,
            ["pentagon"] = (int)MsoAutoShapeType.msoShapePentagon,
            ["hexagon"] = (int)MsoAutoShapeType.msoShapeHexagon,
            ["octagon"] = (int)MsoAutoShapeType.msoShapeOctagon,
            ["pie"] = (int)MsoAutoShapeType.msoShapePie,
            ["chord"] = (int)MsoAutoShapeType.msoShapeChord,
            ["donut"] = (int)MsoAutoShapeType.msoShapeDonut,
            ["foldedCorner"] = (int)MsoAutoShapeType.msoShapeFoldedCorner,
            ["heart"] = (int)MsoAutoShapeType.msoShapeHeart,
            ["lightningBolt"] = (int)MsoAutoShapeType.msoShapeLightningBolt,
            ["sun"] = (int)MsoAutoShapeType.msoShapeSun,
            ["moon"] = (int)MsoAutoShapeType.msoShapeMoon,
            ["cloud"] = (int)MsoAutoShapeType.msoShapeCloud,
            ["arc"] = (int)MsoAutoShapeType.msoShapeArc,
            ["star5"] = (int)MsoAutoShapeType.msoShape5pointStar,
            ["rightArrow"] = (int)MsoAutoShapeType.msoShapeRightArrow,
            ["leftArrow"] = (int)MsoAutoShapeType.msoShapeLeftArrow,
            ["upArrow"] = (int)MsoAutoShapeType.msoShapeUpArrow,
            ["downArrow"] = (int)MsoAutoShapeType.msoShapeDownArrow,
        };
    }
}

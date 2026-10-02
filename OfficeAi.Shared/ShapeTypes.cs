using System;
using System.Collections.Generic;
using MsoAutoShapeType = Microsoft.Office.Core.MsoAutoShapeType;

namespace OfficeAi.Shared
{
    /// <summary>
    /// Shape-name to MsoAutoShapeType, as int (not the enum type) - an
    /// embedded interop type can't be used as a generic type argument across
    /// an assembly boundary (CS1769). Callers cast:
    ///     (Microsoft.Office.Core.MsoAutoShapeType)ShapeTypes.ByName["rect"]
    /// "textbox" is handled separately by each app and is not in this map.
    /// Full background: ShapeTypes.cs.md.
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

# PowerPointTools.Images.cs

## `CropImage`

Approximation, documented deliberately: fractions are applied against
the shape's CURRENT on-slide size, not the original uncropped source
image - classic Interop has no reliable "natural size" property once a
picture has already been resized/cropped on the slide. Correct for a
freshly-inserted, never-before-cropped picture; imprecise under
repeated crop calls on the same shape.

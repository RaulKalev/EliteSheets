using netDxf;
using netDxf.Blocks;
using netDxf.Entities;
using netDxf.Objects;
using netDxf.Tables;
using netDxf.Units;
using System;
using System.IO;

namespace EliteSheets.Services
{
    /// <summary>
    /// Post-processes DXF files: inserts a title block (from a DXF template) into MODEL space.
    /// Assumes both exported DXF and template DXF are in millimetres.
    /// </summary>
    public class DxfPostProcessor
    {
        /// <summary>
        /// Insert a title block DXF into a target DXF at bottom-left, scaled to sheet size * viewScale.
        /// </summary>
        public void InsertTitleBlockIntoModel(
            string targetDxfPath,
            string titleBlockDxfPath,
            double sheetWidthMm,
            double sheetHeightMm,
            double viewScale,
            double templateWidthMm,
            double templateHeightMm,
            double marginMm = 10.0)
        {
            if (!File.Exists(targetDxfPath))
                throw new FileNotFoundException("Target DXF not found.", targetDxfPath);
            if (!File.Exists(titleBlockDxfPath))
                throw new FileNotFoundException("Title-block DXF not found.", titleBlockDxfPath);
            if (templateWidthMm <= 0 || templateHeightMm <= 0)
                throw new ArgumentException("Template size must be positive.");

            // Load docs
            var target = DxfDocument.Load(targetDxfPath);
            var template = DxfDocument.Load(titleBlockDxfPath);

            // Units (recommended: keep everything in mm)
            target.DrawingVariables.InsUnits = DrawingUnits.Millimeters;

            // Compute non-uniform scale to match printed sheet size at current view scale
            // printed_mm = model_mm / viewScale => model_mm = printed_mm * viewScale
            double sx = (sheetWidthMm / templateWidthMm) * viewScale;
            double sy = (sheetHeightMm / templateHeightMm) * viewScale;

            // Anchor in model space with a small paper margin
            double x = marginMm * viewScale;
            double y = marginMm * viewScale;

            // Build a block from TEMPLATE model space
            string blockName = $"TB_{Guid.NewGuid():N}";
            Block tbBlock = new Block(blockName);

            // Find Model Space block in the template
            Block modelBlock = null;
            if (template.Blocks.Contains(Layout.ModelSpaceName))
            {
                modelBlock = template.Blocks[Layout.ModelSpaceName];
            }
            else
            {
                foreach (var b in template.Blocks)
                {
                    var name = b.Name?.Trim();
                    if (string.Equals(name, "*Model_Space", StringComparison.OrdinalIgnoreCase) ||
                        string.Equals(name, "*MODEL_SPACE", StringComparison.OrdinalIgnoreCase))
                    {
                        modelBlock = b;
                        break;
                    }
                }
            }
            if (modelBlock == null)
                throw new InvalidDataException("Model space block not found in title-block DXF.");

            // Copy all model-space entities into our block
            foreach (var ent in modelBlock.Entities)
            {
                var clone = (EntityObject)ent.Clone();
                tbBlock.Entities.Add(clone);
            }

            // Register the block
            target.Blocks.Add(tbBlock);

            // Insert into target MODEL space with position + non-uniform scale
            var ins = new Insert(tbBlock)
            {
                Position = new netDxf.Vector3(x, y, 0),
                Scale = new netDxf.Vector3(sx, sy, 1.0)
            };
            target.Entities.Add(ins);

            // (optional but recommended) lock units to mm
            target.DrawingVariables.InsUnits = DrawingUnits.Millimeters;

            target.Save(targetDxfPath);

        }
    }
}

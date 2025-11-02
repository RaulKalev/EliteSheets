using netDxf;
using netDxf.Blocks;
using netDxf.Entities;
using netDxf.Objects;
using netDxf.Tables;
using netDxf.Units;
using System;
using System.IO;
using System.Linq;

namespace EliteSheets.Services
{
    /// <summary>
    /// Promotes Paper Space contents (title block etc.) and merges the first viewport’s model geometry into Model Space.
    /// Uses a transform derived from the first paper-space viewport (ViewHeight / Height).
    /// </summary>
    public class DxfPaperToModelPromoter
    {
        /// <param name="contentOnly">
        /// If true, does NOT copy paper-space entities (title block) into Model space.
        /// Only the model geometry (aligned by the first viewport) is inserted.
        /// </param>
        public void PromotePaperToModel(string dxfPath, bool contentOnly = false)
        {
            if (!File.Exists(dxfPath))
                throw new FileNotFoundException("DXF not found.", dxfPath);

            var doc = DxfDocument.Load(dxfPath);
            doc.DrawingVariables.InsUnits = DrawingUnits.Millimeters;

            var paperLayout = doc.Layouts
                .FirstOrDefault(l => !string.Equals(l.Name, Layout.ModelSpaceName, StringComparison.OrdinalIgnoreCase));

            if (paperLayout == null || paperLayout.AssociatedBlock == null)
            {
                // No paper at all -> nothing to do
                return;
            }

            var paperBlock = paperLayout.AssociatedBlock;

            Block modelBlock = doc.Blocks.Contains(Layout.ModelSpaceName)
                ? doc.Blocks[Layout.ModelSpaceName]
                : doc.Blocks.FirstOrDefault(b => b.Name.Equals("*Model_Space", StringComparison.OrdinalIgnoreCase));

            // If there’s no viewport, we can’t align; for contentOnly, just bail (keep as-is).
            var vp = paperBlock.Entities.OfType<Viewport>().FirstOrDefault();
            if (vp == null)
            {
                if (!contentOnly)
                {
                    // Old behavior: copy paper only, so at least title exists in model space.
                    if (paperBlock.Entities.Count > 0)
                    {
                        var tb = new Block($"PS_ONLY_{Guid.NewGuid():N}");
                        foreach (var e in paperBlock.Entities)
                            if (!(e is Viewport)) tb.Entities.Add((EntityObject)e.Clone());
                        doc.Blocks.Add(tb);
                        doc.Entities.Add(new Insert(tb) { Position = new netDxf.Vector3(0, 0, 0) });
                        doc.Save(dxfPath);
                    }
                }
                return;
            }

            // Compute transform to align model to the viewport center/scale
            double vpScale = (vp.Height > 1e-9) ? (vp.ViewHeight / vp.Height) : 1.0;
            double paperX = vp.Center.X, paperY = vp.Center.Y;
            double modelX = vp.ViewCenter.X, modelY = vp.ViewCenter.Y;
            double scale = 1.0 / vpScale, dx = paperX - (modelX / vpScale), dy = paperY - (modelY / vpScale);

            var combined = new Block($"PROMOTED_{Guid.NewGuid():N}");

            // NEW: add paper entities only if contentOnly == false
            if (!contentOnly)
            {
                foreach (var ent in paperBlock.Entities)
                    if (!(ent is Viewport)) combined.Entities.Add((EntityObject)ent.Clone());
            }

            if (modelBlock != null)
            {
                var modelCopy = new Block($"MODEL_COPY_{Guid.NewGuid():N}");
                foreach (var ent in modelBlock.Entities)
                    modelCopy.Entities.Add((EntityObject)ent.Clone());
                doc.Blocks.Add(modelCopy);

                combined.Entities.Add(new Insert(modelCopy)
                {
                    Position = new netDxf.Vector3(dx, dy, 0),
                    Scale = new netDxf.Vector3(scale, scale, 1),
                    Rotation = 0
                });
            }

            // Register the combined block into the CURRENT doc
            doc.Blocks.Add(combined);

            // ---- Build a fresh document that only contains our combined insert ----
            var newDoc = new DxfDocument();
            newDoc.DrawingVariables.InsUnits = doc.DrawingVariables.InsUnits;

            // Copy the block definition into the new doc
            var combinedClone = new Block(combined.Name);
            foreach (var ent in combined.Entities)
                combinedClone.Entities.Add((EntityObject)ent.Clone());
            newDoc.Blocks.Add(combinedClone);

            // Place one instance at origin
            newDoc.Entities.Add(new Insert(combinedClone)
            {
                Position = new netDxf.Vector3(0, 0, 0),
                Scale = new netDxf.Vector3(1, 1, 1),
                Rotation = 0
            });

            // Overwrite original file
            newDoc.Save(dxfPath);

        }

        private static Block FindPaperSpaceBlock(DxfDocument doc)
        {
            // Prefer the first non-Model layout (paper) and take its AssociatedBlock
            var paperLayout = doc.Layouts.FirstOrDefault(l =>
                !string.Equals(l.Name, Layout.ModelSpaceName, StringComparison.OrdinalIgnoreCase));

            if (paperLayout?.AssociatedBlock != null)
                return paperLayout.AssociatedBlock;

            // Fallback to common paper-space block names
            return doc.Blocks.FirstOrDefault(b =>
                string.Equals(b.Name, "*Paper_Space", StringComparison.OrdinalIgnoreCase) ||
                string.Equals(b.Name, "*PAPER_SPACE", StringComparison.OrdinalIgnoreCase) ||
                string.Equals(b.Name, "*Paper_Space0", StringComparison.OrdinalIgnoreCase) ||
                string.Equals(b.Name, "*PAPER_SPACE0", StringComparison.OrdinalIgnoreCase));
        }
    }
}

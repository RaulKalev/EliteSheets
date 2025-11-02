using netDxf;
using netDxf.Blocks;
using netDxf.Entities;
using netDxf.Objects;
using netDxf.Units;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;

namespace EliteSheets.Services
{
    public class DxfMergeService
    {
        /// <summary>
        /// Merge the given DXF files into a single DXF by inserting the main content block
        /// (e.g., PROMOTED_...) from each file, stacked vertically (mm units).
        /// </summary>
        public void MergeIntoSingleDxf(
            IList<string> sourceDxfPaths,
            string outputDxfPath,
            double sheetSpacingMm = 220.0)
        {
            if (sourceDxfPaths == null || sourceDxfPaths.Count == 0)
                throw new ArgumentException("No source DXF files to merge.", nameof(sourceDxfPaths));

            Directory.CreateDirectory(Path.GetDirectoryName(outputDxfPath) ?? ".");

            var target = new DxfDocument();
            target.DrawingVariables.InsUnits = DrawingUnits.Millimeters;

            double currentX = 0.0;
            int idx = 0;

            foreach (var path in sourceDxfPaths)
            {
                if (string.IsNullOrWhiteSpace(path) || !File.Exists(path))
                    continue;

                var src = DxfDocument.Load(path);

                // 1) Try to find a “content block” created by our promoter (PROMOTED_/PS_ONLY_/MODEL_COPY_…)
                //    We simply pick the first non-reserved block with any entities.
                Block content = src.Blocks
                    .FirstOrDefault(b =>
                        !string.IsNullOrEmpty(b.Name) &&
                        b.Name[0] != '*' &&                              // skip *Model_Space, *Paper_Space, etc.
                        b.Entities != null && b.Entities.Count > 0);

                // 2) Fallback: if no suitable block, try the Model Space associated block
                if (content == null)
                {
                    var modelLayout = src.Layouts.FirstOrDefault(l => string.Equals(l.Name, Layout.ModelSpaceName, StringComparison.OrdinalIgnoreCase));
                    content = modelLayout?.AssociatedBlock;
                    if (content != null && (content.Entities == null || content.Entities.Count == 0))
                        content = null;
                }

                // 3) If still nothing, skip this file
                if (content == null)
                    continue;

                // 4) Clone that block into TARGET with a unique name
                string mergedBlockName = $"MERGE_{idx}_{Guid.NewGuid():N}";
                var cloned = new Block(mergedBlockName);
                foreach (var ent in content.Entities)
                    cloned.Entities.Add((EntityObject)ent.Clone());

                target.Blocks.Add(cloned);

                // 5) Insert it in Model space with vertical offset so sheets don’t overlap
                var ins = new Insert(cloned)
                {
                    Position = new netDxf.Vector3(currentX, 0, 0),   // ⟵ X offset
                    Scale = new netDxf.Vector3(1, 1, 1),
                    Rotation = 0
                };
                target.Entities.Add(ins);

                currentX += sheetSpacingMm;
                idx++;
            }

            target.Save(outputDxfPath);
        }
    }
}

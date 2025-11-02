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
        // ... keep your existing methods ...

        /// <summary>
        /// Merges multiple DXFs side-by-side into a temporary in-memory document (model-space only),
        /// then inserts that merged content into the given TEMPLATE DXF's Model Space and saves to output.
        /// This avoids bringing any layouts/viewports from sources into the template.
        /// </summary>
        public void MergeIntoTemplate(
            IList<string> sourceDxfPaths,
            string templateDxfPath,
            string outputDxfPath,
            double sheetSpacingMm = 220.0,
            double insertXmm = 0.0,
            double insertYmm = 0.0)
        {
            if (sourceDxfPaths == null || sourceDxfPaths.Count == 0)
                throw new ArgumentException("No source DXF files to merge.", nameof(sourceDxfPaths));
            if (string.IsNullOrWhiteSpace(templateDxfPath) || !File.Exists(templateDxfPath))
                throw new FileNotFoundException("Template DXF not found.", templateDxfPath);

            Directory.CreateDirectory(Path.GetDirectoryName(outputDxfPath) ?? ".");

            // 1) Build an in-memory "merged" doc with only model-space inserts
            var merged = new DxfDocument();
            merged.DrawingVariables.InsUnits = DrawingUnits.Millimeters;

            // Keep track of the inserts we add (since DrawingEntities isn't enumerable)
            var mergedInserts = new List<Insert>();

            double currentX = 0.0;
            int idx = 0;

            foreach (var path in sourceDxfPaths)
            {
                if (string.IsNullOrWhiteSpace(path) || !File.Exists(path))
                    continue;

                var src = DxfDocument.Load(path);

                // Prefer first non-reserved block with entities; else fallback to Model space block
                Block content = src.Blocks
                    .FirstOrDefault(b => !string.IsNullOrEmpty(b.Name)
                                         && b.Name[0] != '*'
                                         && b.Entities != null
                                         && b.Entities.Count > 0);

                if (content == null)
                {
                    var modelLayout = src.Layouts
                        .FirstOrDefault(l => string.Equals(l.Name, Layout.ModelSpaceName, StringComparison.OrdinalIgnoreCase));
                    content = modelLayout?.AssociatedBlock;
                    if (content != null && (content.Entities == null || content.Entities.Count == 0))
                        content = null;
                }
                if (content == null)
                    continue;

                // Clone content into a unique block within the MERGED doc
                string mergedBlockName = $"MERGE_{idx}_{Guid.NewGuid():N}";
                var cloned = new Block(mergedBlockName);
                foreach (var e in content.Entities)
                    cloned.Entities.Add((EntityObject)e.Clone());

                merged.Blocks.Add(cloned);

                // Insert this "page" at currentX offset
                var ins = new Insert(cloned)
                {
                    Position = new netDxf.Vector3(currentX, 0, 0),
                    Scale = new netDxf.Vector3(1, 1, 1),
                    Rotation = 0
                };
                merged.Entities.Add(ins);
                mergedInserts.Add(ins);

                currentX += sheetSpacingMm;
                idx++;
            }

            // Safety: strip any paper layouts in the merged doc (we won't copy them anyway)
            // (This loop *is* enumerable in netDxf)
            foreach (var layout in merged.Layouts.ToList())
            {
                if (!layout.Name.Equals(Layout.ModelSpaceName, StringComparison.OrdinalIgnoreCase))
                    merged.Layouts.Remove(layout.Name);
            }

            // 2) Load TEMPLATE and copy merged content INTO template's Model Space
            var template = DxfDocument.Load(templateDxfPath);
            template.DrawingVariables.InsUnits = DrawingUnits.Millimeters;

            // Clone blocks from merged -> template, tracking names to handle collisions
            var blockNameMap = new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase);

            foreach (var b in merged.Blocks)
            {
                if (string.IsNullOrEmpty(b.Name) || b.Name[0] == '*') continue;

                string newName = b.Name;
                if (template.Blocks.Contains(newName))
                {
                    // Generate a collision-free name (keep it short-ish to avoid DXF viewer quirks)
                    newName = $"_{Guid.NewGuid():N}".Substring(0, 12);
                }

                var cloneBlock = new Block(newName);
                foreach (var ent in b.Entities)
                    cloneBlock.Entities.Add((EntityObject)ent.Clone());

                template.Blocks.Add(cloneBlock);
                blockNameMap[b.Name] = newName;
            }

            // Insert all MERGED model-space inserts into TEMPLATE model space
            double dx = insertXmm;
            double dy = insertYmm;

            foreach (var ins in mergedInserts)
            {
                string refName = ins.Block?.Name;
                if (string.IsNullOrEmpty(refName)) continue;

                // Map to cloned name in template
                if (!blockNameMap.TryGetValue(refName, out string newRefName))
                    continue;

                if (!template.Blocks.Contains(newRefName))
                    continue;

                var blockInTemplate = template.Blocks[newRefName];

                var insClone = new Insert(blockInTemplate)
                {
                    Position = new netDxf.Vector3(ins.Position.X + dx, ins.Position.Y + dy, 0),
                    Scale = new netDxf.Vector3(ins.Scale.X, ins.Scale.Y, ins.Scale.Z),
                    Rotation = ins.Rotation
                };
                template.Entities.Add(insClone);
            }
            // --- Cleanup unwanted Paper Space entities (keep locked layer content) ---
            foreach (var layout in template.Layouts)
            {
                if (layout.Name.Equals(Layout.ModelSpaceName, StringComparison.OrdinalIgnoreCase))
                    continue; // skip model space

                var layoutBlock = layout.AssociatedBlock;
                if (layoutBlock == null) continue;

                // Collect entities to delete
                var toRemove = new List<EntityObject>();
                foreach (var ent in layoutBlock.Entities)
                {
                    var layer = ent.Layer;
                    bool isLocked = layer != null && layer.IsLocked;

                    // If it's a viewport or non-locked content, mark for deletion
                    if (!isLocked && ent.Type == EntityType.Viewport)
                    {
                        toRemove.Add(ent);
                    }
                    else if (!isLocked && ent.Type != EntityType.Viewport)
                    {
                        // optionally: clear other stray stuff not on locked layer
                        toRemove.Add(ent);
                    }
                }

                // Delete all marked entities
                foreach (var ent in toRemove)
                    layoutBlock.Entities.Remove(ent);
            }

            // 3) Save final file (preserving template's layouts/viewports untouched)
            template.Save(outputDxfPath);
        }

    }
}

using Autodesk.Revit.DB;
using System;
using System.Collections.Generic;
using System.Linq;

namespace EliteSheets.Helpers
{
    public static class PaperSizeHelper
    {
        private const double Tolerance = 10.0; // millimeters

        // ISO paper sizes in millimeters (width x height)
        private static readonly Dictionary<string, (int Width, int Height)> IsoSizes = new Dictionary<string, (int, int)>
        {
            { "A0", (841, 1189) },
            { "A1", (594, 841) },
            { "A2", (420, 594) },
            { "A3", (297, 420) },
            { "A4", (210, 297) },
            { "A5", (148, 210) },
            { "A6", (105, 148) }
        };

        /// <summary>
        /// Returns the ISO paper label (e.g., A3) or custom size like "275x390mm"
        /// </summary>
        public static string GetPaperSizeLabel(ViewSheet sheet)
        {
            if (sheet == null) return "Unknown";

            var outline = sheet.Outline;
            double widthMm = UnitUtils.ConvertFromInternalUnits(outline.Max.U - outline.Min.U, UnitTypeId.Millimeters);
            double heightMm = UnitUtils.ConvertFromInternalUnits(outline.Max.V - outline.Min.V, UnitTypeId.Millimeters);

            double w = Math.Min(widthMm, heightMm);
            double h = Math.Max(widthMm, heightMm);

            // Try to match ISO label
            foreach (var kvp in IsoSizes)
            {
                if (Math.Abs(w - kvp.Value.Width) <= Tolerance && Math.Abs(h - kvp.Value.Height) <= Tolerance)
                {
                    return kvp.Key;
                }
            }

            // Return as custom size
            return $"{Math.Round(widthMm)}x{Math.Round(heightMm)}mm";
        }

        /// <summary>
        /// Gets the width and height of the sheet in millimeters.
        /// </summary>
        public static (double WidthMm, double HeightMm) GetSizeMm(ViewSheet sheet)
        {
            if (sheet == null)
                return (0, 0);

            var outline = sheet.Outline;
            double width = UnitUtils.ConvertFromInternalUnits(outline.Max.U - outline.Min.U, UnitTypeId.Millimeters);
            double height = UnitUtils.ConvertFromInternalUnits(outline.Max.V - outline.Min.V, UnitTypeId.Millimeters);

            return (width, height);
        }

        public static bool IsStandardSize(string label)
        {
            return IsoSizes.ContainsKey(label);
        }
        public static string GetPaperSizeLabel(double widthMm, double heightMm)
        {
            double w = Math.Min(widthMm, heightMm);
            double h = Math.Max(widthMm, heightMm);
            const double Tolerance = 10.0;

            var isoSizes = new Dictionary<string, (int Width, int Height)>
    {
        { "A0", (841, 1189) },
        { "A1", (594, 841) },
        { "A2", (420, 594) },
        { "A3", (297, 420) },
        { "A4", (210, 297) },
        { "A5", (148, 210) },
        { "A6", (105, 148) }
    };

            foreach (var size in isoSizes)
            {
                if (Math.Abs(w - size.Value.Width) <= Tolerance && Math.Abs(h - size.Value.Height) <= Tolerance)
                {
                    return size.Key;
                }
            }

            return $"{Math.Round(widthMm)} x {Math.Round(heightMm)} mm";
        }

    }
}

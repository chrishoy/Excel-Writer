namespace ExportMap.TestHarness
{
    using System.Collections.Generic;

    using GamFX.Domain.Report.PositionReport;

    internal static class PositionReportGroupBuilder
    {
        public static PositionReportGroup Build()
        {
            var group = new PositionReportGroup();
            group.Rows = new List<PositionReportRow>();
            return group;
        }

        public static PositionReportGroup WithHeading(this PositionReportGroup source, string value)
        {
            source.Heading = value;
            return source;
        }

        public static PositionReportGroup WithRow(this PositionReportGroup source, PositionReportRow value)
        {
            if (source.Rows == null)
            {
                source.Rows = new List<PositionReportRow> { value };
            }
            else
            {
                source.Rows.Add(value);
            }

            return source;
        }
    }
}
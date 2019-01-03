namespace ExportMap.TestHarness
{
    using System.Collections.Generic;

    using GamFX.Domain.Report.PositionReport;

    internal static class PositionReportBuilder
    {
        public static PositionReport Build()
        {
            var report = new PositionReport
                             {
                                 SubHeadings = new List<string>(),
                                 ValidationErrors = new List<string>(),
                                 Groups = new List<PositionReportGroup>(),
                             };
            return report;
        }

        public static PositionReport WithHeading(this PositionReport source, string value)
        {
            source.Heading = value;
            return source;
        }

        public static PositionReport WithSubHeading(this PositionReport source, string value)
        {
            if (source.SubHeadings == null)
            {
                source.SubHeadings = new List<string> { value };
            }
            else
            {
                source.SubHeadings.Add(value);
            }

            return source;
        }

        public static PositionReport WithValidationError(this PositionReport source, string value)
        {
            if (source.ValidationErrors == null)
            {
                source.ValidationErrors = new List<string> { value };
            }
            else
            {
                source.ValidationErrors.Add(value);
            }

            return source;
        }

        public static PositionReport WithGroup(this PositionReport source, PositionReportGroup value)
        {
            if (source.Groups == null)
            {
                source.Groups = new List<PositionReportGroup> { value };
            }
            else
            {
                source.Groups.Add(value);
            }

            return source;
        }
    }
}
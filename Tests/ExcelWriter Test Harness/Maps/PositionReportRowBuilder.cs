namespace ExportMap.TestHarness
{
    using System;

    using GamFX.Domain.Fund;
    using GamFX.Domain.Report.PositionReport;
    using GamFX.Domain.Transaction;

    internal static class PositionReportRowBuilder
    {
        public static PositionReportRow Build()
        {
            var row = new PositionReportRow();
            return row;
        }

        public static PositionReportRow WithFund(this PositionReportRow source, Fund value)
        {
            source.Fund = value;
            return source;
        }

        public static PositionReportRow WithHedgeCurrency(this PositionReportRow source, string value)
        {
            source.HedgeCurrency = value;
            return source;
        }

        public static PositionReportRow WithFundValue(this PositionReportRow source, decimal? value)
        {
            source.FundValue = value;
            return source;
        }

        public static PositionReportRow WithFundValueDate(this PositionReportRow source, DateTime? value)
        {
            source.FundValueDate = value;
            return source;
        }

        public static PositionReportRow WithCoverRequired(this PositionReportRow source, decimal value)
        {
            source.CoverRequired = value;
            return source;
        }

        public static PositionReportRow WithCurrentLevel(this PositionReportRow source, decimal value)
        {
            source.CurrentLevel = value;
            return source;
        }

        public static PositionReportRow WithDealing(this PositionReportRow source, decimal value)
        {
            source.Dealing = value;
            return source;
        }

        public static PositionReportRow WithForwardPosition(this PositionReportRow source, decimal value)
        {
            source.ForwardPosition = value;
            return source;
        }

        public static PositionReportRow WithStatus(this PositionReportRow source, PositionStatus value)
        {
            source.Status = value;
            return source;
        }

        public static PositionReportRow WithDate(this PositionReportRow source, DateTime value)
        {
            source.LastDealingDate = value;
            return source;
        }

        public static PositionReportRow WithHighlight(this PositionReportRow source, bool value)
        {
            source.Highlight = value;
            return source;
        }
        public static PositionReportRow WithMessage(this PositionReportRow source, string value)
        {
            source.Message = value;
            return source;
        }
    }
}
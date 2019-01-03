namespace GamFX.Business.FxReports
{
    using System;
    using System.Collections.Generic;
    using System.Diagnostics;
    using System.Linq;

    using FOF.Framework;
    using FOF.Framework.ComponentModel.Collections;
    using FOF.Infrastructure.Logging.Interface;

    using GamFX.Data;
    using GamFX.Domain.Fund;
    using GamFX.Domain.Mfgi;
    using GamFX.Domain.Report.PositionReport;
    using GamFX.Domain.Transaction;
    using GamFX.ServiceInfrastructure.Identity;
    using GamFX.ServiceInfrastructure.Services;

    /// <summary>
    /// Command which, when executed, gets the data for the generation of a position report.
    /// </summary>
    public class GetPositionReportDataCommand
    {
        private readonly ILogger logger;
        private readonly IUserIdentity userIdentity;
        private readonly IDateTimeService dateTimeService;
        private readonly IFundRepository fundRepository;
        private readonly IValuationRepository valuationRepository;
        private readonly IOrderManagementService orderManagementService;
        private readonly IForwardContractRepository forwardContractRepository;
        private readonly IMfgiRepository mfgiRepository;
        private readonly IStructureRepository structureRepository;

        /// <summary>
        /// Initialises a new instance of the <see cref="GetPositionReportDataCommand"/> class.
        /// </summary>
        /// <param name="logger">The logger.</param>
        /// <param name="userIdentity">The user identity.</param>
        /// <param name="dateTimeService">The date-time service.</param>
        /// <param name="fundRepository">The fund repository.</param>
        /// <param name="valuationRepository">The valuation repository</param>
        /// <param name="orderManagementService">The order management service</param>
        /// <param name="forwardContractRepository">The forward Contract Repository</param>
        /// <param name="mfgiRepository">The MFGI Repository.</param>
        /// <param name="structureRepository">The fund structure repository</param>
        public GetPositionReportDataCommand(
            ILogger logger,
            IUserIdentity userIdentity,
            IDateTimeService dateTimeService,
            IFundRepository fundRepository,
            IValuationRepository valuationRepository,
            IOrderManagementService orderManagementService,
            IForwardContractRepository forwardContractRepository,
            IMfgiRepository mfgiRepository,
            IStructureRepository structureRepository)

        {
            Guard.IsNotNull(logger, "logger");
            Guard.IsNotNull(userIdentity, "uderIdentity");
            Guard.IsNotNull(dateTimeService, "dateTimeService");
            Guard.IsNotNull(fundRepository, "fundRepository");
            Guard.IsNotNull(valuationRepository, "valuationRepository");
            Guard.IsNotNull(orderManagementService, "orderManagementService");
            Guard.IsNotNull(forwardContractRepository, "forwardContractRepository");
            Guard.IsNotNull(mfgiRepository, "mfgiRepository");
            Guard.IsNotNull(structureRepository, "structureRepository");

            this.logger = logger;
            this.userIdentity = userIdentity;
            this.dateTimeService = dateTimeService;
            this.fundRepository = fundRepository;
            this.valuationRepository = valuationRepository;
            this.orderManagementService = orderManagementService;
            this.forwardContractRepository = forwardContractRepository;
            this.mfgiRepository = mfgiRepository;
            this.structureRepository = structureRepository;
        }

        /// <summary>
        /// Executes the command.
        /// </summary>
        /// <param name="reportParameters">Parameters used to get the data.</param>
        /// <returns></returns>
        public PositionReport Execute(PositionReportParameters reportParameters)
        {
            this.logger.Log(LogType.Trace, this.GetAssemblyName(), "Enter {0}.Execute", this.GetType().Name);

            Guard.IsNotNull(reportParameters, "reportParameters");

            var data = this.BuildData(reportParameters);

            this.logger.Log(LogType.Trace, this.GetAssemblyName(), "Exit {0}.Execute", this.GetType().Name);

            return data;
        }

        /// <summary>
        /// Gets the report data in cases where the data has not been supplied
        /// </summary>
        /// <param name="reportParameters">The report parameters.</param>
        /// <returns>Report data where possible, or null if not.</returns>
        public PositionReport BuildData(PositionReportParameters reportParameters)
        {
            var user = this.userIdentity.Username;
            var runDateTime = this.dateTimeService.Now;

            var date = runDateTime.Date;
            string fundGroups = reportParameters.FundGroups.Any()
                                     ? string.Join(",", reportParameters.FundGroups.ConvertAll(i => i.Code))
                                     : "ALL";

            var fundGroupIds = reportParameters.FundGroups.Any()
                                   ? reportParameters.FundGroups.Select(g => g.FundGroupId).ToList()
                                   : null;

            var fundsWithDealingCalendars = this.GetRequiredFundsWithDealingCalendars(fundGroupIds, date).ToList();
            var groups = fundsWithDealingCalendars.ConvertAll(f => f.Fund.FundGroup).Distinct();

            // Get current values for funds selected (at the value date specified by its dealing calndar (or null for latest))
            var valuationParams =
                fundsWithDealingCalendars.ConvertAll(
                    f =>
                    new ValuationParameters
                        {
                            CdbPrtId = f.Fund.PrtId,
                            HedgeType = f.Fund.HedgeType,
                            ValuationType = f.Fund.ValuationType,
                            ValueDate = f.DealingCalendar == null ? null : f.DealingCalendar.DealingDate, // No dealing calendar = get latest valuation
                        });
            var valuations = this.valuationRepository.GetValuationsByCdbPrtIdList(valuationParams);

            // Get core trades for funds selected
            //var tradeBlotterOptions = new CoreBlotterOptions { };
            //this.orderManagementService.GetCoreTrades()

            var forwardContracts = forwardContractRepository.GetForwardContractsByFundIdList(fundsWithDealingCalendars.ConvertAll(f => f.Fund.FundId));

            var r = new Random();
            var report =
                PositionReportBuilder.Build()
                    .WithHeading("Position Report")
                    .WithSubHeading(string.Format("Daily Hedge Position Report for {0} funds.", fundGroups))
                    .WithSubHeading(string.Format("Run on {0:dd'-'MMM'-'yyyy} at {0:hh:mm:ss} by {1}", runDateTime, user))
                    .WithValidationError("This is a message that could be generated and shown at the top of the report! For example:");

            groups.ForEach(
                g =>
                    {
                        var group = PositionReportGroupBuilder.Build().WithHeading(g.Name);
                        report.WithGroup(group);

                        (from fundWithDealingCalendar in fundsWithDealingCalendars
                         where fundWithDealingCalendar.Fund.FundGroup.Equals(g)
                         select fundWithDealingCalendar.Fund).ForEach(
                            fnd =>
                                {
                                    Valuation val = valuations.SingleOrDefault(v => v.CdbPrtId == fnd.PrtId);
                                    ForwardContract fwd = forwardContracts.SingleOrDefault(fw => fw.FundId == fnd.FundId);

                                    var randomError = r.Next(1, 100);
                                    var row =
                                        PositionReportRowBuilder.Build()
                                            .WithFund(fnd)
                                            .WithFundValue(val == null ? null : val.FundValue)
                                            .WithFundValueDate(val == null ? null : val.ValueDate)
                                            .WithHedgeCurrency(fnd.HedgeType == HedgeType.Investment ? fnd.FeedsInto.Ccy : fnd.Ccy)
                                            .WithCurrentLevel(GetCurrentLevel(fwd, fnd.HedgeType));

                                    // Example error
                                    if (randomError < 6)
                                    {
                                        row.WithHighlight(true).WithMessage("Error " + randomError);
                                        report.WithValidationError(string.Format("{0} has {1}", row.Fund.GamFundCode, row.Message));
                                    }

                                    group.WithRow(row);
                                });

                    });

            return report;
        }

        private IEnumerable<FundWithDealingCalendar> GetRequiredFundsWithDealingCalendars(List<decimal> fundGroupIds, DateTime date)
        {
            // First, get MFGI dealing calendars (currently for all as it's MFGI based) for all funds that will price on the specified value date
            var prtIdsWithdealingCalendars = this.GetPrtIdsWithDealingCalendarsOfFundsThatArePricing(date);

            // De-duplicate prt-id's getting first (1 prtid can map to 2 MFGI funds)
            var prtIdsWithdealingCalendars2 = (from p in prtIdsWithdealingCalendars
                                               group p by p.Key
                                               into g select g.First()).ToList();

            var prtIdList = prtIdsWithdealingCalendars2.ConvertAll(i => i.Key);

            // Then lookup funds, restricting to hedged and in the required groups (From Rene, we only get funds in groups that are pricing)
            var funds = this.fundRepository.FindFunds(new FundQuery { IsHedged = true, FundGroupIds = fundGroupIds, PrtIds = prtIdList });

            // Join funds back to dealing calendars (so we have value dates that are pricing)
            var fundsWithDealingCalendars = from fund in funds
                                            join prtIdWithdealingCalendar in prtIdsWithdealingCalendars2
                                                on fund.PrtId equals prtIdWithdealingCalendar.Key
                                            select new FundWithDealingCalendar(fund, prtIdWithdealingCalendar.Value);

            return fundsWithDealingCalendars;
        }

        private IEnumerable<KeyValuePair<decimal, DealingCalendar>> GetPrtIdsWithDealingCalendarsOfFundsThatArePricing(DateTime date)
        {
            // Get list of ALL fund structures - At the root will be core funds
            var coreStructures = this.structureRepository.GetStructureTree();

            // Get price dates (i.e. when PR2 will run on both classes and cores) for a value date
            var allMfgiPriceDates = this.mfgiRepository.GetFundsToPrice(date);

            // Get all feeders for all cores 
            var feederStructures = coreStructures.SelectMany(c => c.Feeders);

            // Match up feeders with price using MFGI TA Fund Code/Share Class (if no match, DealingCalendar will be null)
            var dealingCalendars = from feederStructure in feederStructures
                                   join dealCal in allMfgiPriceDates on
                                       new { feederStructure.TaFundCode, feederStructure.ShareClass } equals
                                       new { dealCal.TaFundCode, dealCal.ShareClass }
                                   select new KeyValuePair<decimal, DealingCalendar>(feederStructure.CdbPrtId, dealCal);

            return dealingCalendars;
        }

        private static decimal GetCurrentLevel(ForwardContract contract, HedgeType hedgeType)
        {
            return contract == null ? 0M : hedgeType == HedgeType.Investment ? -contract.Cover : contract.Cover;
        }

        [DebuggerDisplay("{DebuggerDisplay,nq}")]
        private class FundWithDealingCalendar
        {
            public FundWithDealingCalendar(Fund fund, DealingCalendar dealingCalendar)
            {
                this.Fund = fund;
                this.DealingCalendar = dealingCalendar;
            }

            public Fund Fund { get; private set; }

            public DealingCalendar DealingCalendar { get; private set; }

            private string DebuggerDisplay
            {
                get
                {
                    return string.Format(
                        "{0}-{1}",
                        this.Fund == null ? "(no fund)" : this.Fund.GamFundCode,
                        DebuggerDisplayDealingCalendar(this.DealingCalendar));
                }
            }

            private static string DebuggerDisplayDealingCalendar(DealingCalendar dc)
            {
                return dc == null
                           ? "(no dealing calendar)"
                           : string.Format(
                               "MFGI={0}-{1},DealingDate={2:g},FundId={3}",
                               dc.TaFundCode,
                               dc.ShareClass,
                               dc.DealingDate,
                               dc.FundId);
            }
        }
    }
}
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace ExcelWriter.TestHarness.Maps.Data
{
    public static class SampleDataBuilder
    {
        public static object BuildStackedBarChartData()
        {
            return new
            {
                FundData = new List<object>
                {
                    BuildStackedBarChartRow(1, "ABC", 100),
                    BuildStackedBarChartRow(2, "DEF", 200),
                    BuildStackedBarChartRow(3, "GHI", -100),
                    BuildStackedBarChartRow(4, "JKL", -100),
                    BuildStackedBarChartRow(5, "MNO", 150),
                }
            };
        }

        public static object BuildStackedBarChartRow(decimal seriesColour, string fundCode, decimal value)
        {
            return new {SeriesColour = seriesColour, FundCode = fundCode, Value = value};
        }

        public static object Build()
        {
            var data = new 
            {
                Heading = "Report Heading",

                SubHeadings = new List<object>
                {
                    "First Sub-Heading",
                    "Second Sub-Heading",
                },

                Data = new List<object>
                {
                    new {X=10, Y=100},
                    new {X=20, Y=200},
                    new {X=30, Y=300},
                    new {X=40, Y=400},
                },

                Groups = new List<object>
                {
                    new 
                    { 
                        GroupHeading = "Group 1", 
                        Rows = new List<object> 
                        { 
                            BuildRow("AAA"), 
                            BuildRow("BBB"), 
                        },
                    },
                    new
                    {
                        GroupHeading = "Group 2",
                        Rows = new List<object> 
                        { 
                            BuildRow("CCC"), 
                            BuildRow("DDD"), 
                        },
                    },
                    new
                    {
                        GroupHeading = "Group 3",
                        Rows = new List<object> 
                        { 
                            BuildRow("EEE"), 
                            BuildRow("FFF"), 
                        },

                    },
                },
            };

            return data;
        }

        private static object BuildRow(string code)
        {
            return new {Code = code, Message = string.Format("Message for {0}", code)};
        }
    }
}

using System.Collections.Generic;
using System.Linq;
using System.Threading.Tasks;
using ExcelConvert.Infrastructure;
using OfficeOpenXml;

namespace ExcelConvert.Converters
{
    public class QuickBookBalanceSheetConverter : IBatchService
    {
        private readonly AppArgs _appArgs;

        public QuickBookBalanceSheetConverter(
            AppArgs appArgs
            )
        {
            _appArgs = appArgs;
        }

        public async Task RunAsync()
        {
            await Task.CompletedTask;

            if (!System.IO.File.Exists(_appArgs.InFileName))
            {
                throw new System.IO.FileNotFoundException($"In file Not Found: {_appArgs.InFileName}");
            }

            var balanceGroups = Parse();

            WriteOutput(balanceGroups);
        }

        private List<BalanceGroup> Parse()
        {
            var groups = new List<BalanceGroup>();
            var row = 0;

            using (var excel = new ExcelPackage(new System.IO.FileInfo(_appArgs.InFileName)))
            {
                var ws = excel.Workbook.Worksheets.First();

                BalanceGroup group = null!;

                //parse records
                for (row = 2; row <= ws.Dimension.End.Row; row++)
                {
                    var r = Record.FromExcel(ws.Cells, row);

                    if (r.IsEmpty)
                    {
                        continue;
                    }

                    if (r.IsGroupHeader)
                    {
                        group = BalanceGroup.FromRecord(r);
                        groups.Add(group);
                        continue;
                    }

                    if (r.IsBalanceRow)
                    {
                        group.Records.Add(BalanceRecord.FromRecord(r));
                        continue;
                    }

                    if (r.IsTotalRow)
                    {
                        group.Total = BalanceRecord.FromRecord(r);
                        continue;
                    }
                }
            }
            return groups;
        }

        private void WriteOutput(List<BalanceGroup> balanceGroups)
        {
            using (var excel = new ExcelPackage(new System.IO.FileInfo(_appArgs.InFileName)))
            {
                if (excel.Workbook.Worksheets.Any(x => x.Name == "Converted"))
                {
                    var sheet = excel.Workbook.Worksheets.FirstOrDefault(x => x.Name == "Converted");
                    excel.Workbook.Worksheets.Delete(sheet);
                }

                var ws = excel.Workbook.Worksheets.Add("Converted");
                var cells = ws.Cells;
                var row = 1;
                var col = 2;

                //Headers
                cells[row, col].Value = BalanceNames.Commitments;
                cells[row, col += 2].Value = BalanceNames.Contributions;
                cells[row, col += 2].Value = BalanceNames.UnfundedCommitments;
                cells[row, col += 2].Value = BalanceNames.Distributions;
                cells[row, col += 2].Value = "Total Ending Balance";

                row++;
                col = 1;
                for (var i = 1; i < 6; i++)
                {
                    cells[row, ++col].Value = "MTD";
                    cells[row, ++col].Value = "LTD";
                }

                //contents
                foreach (var balanceGroup in balanceGroups.OrderBy(x => x.GroupName))
                {
                    var commitmentsRow = balanceGroup.Records.FirstOrDefault(x => x.Name == BalanceNames.Commitments) ?? new BalanceRecord();
                    var contributionsRow = balanceGroup.Records.FirstOrDefault(x => x.Name == BalanceNames.Contributions) ?? new BalanceRecord();
                    var unfundedCommitmentsRow = balanceGroup.Records.FirstOrDefault(x => x.Name == BalanceNames.UnfundedCommitments) ?? new BalanceRecord();
                    var distributionsRow = balanceGroup.Records.FirstOrDefault(x => x.Name == BalanceNames.Distributions) ?? new BalanceRecord();
                    var totalEndingBalanceRow = balanceGroup.Total ?? new BalanceRecord();

                    row++;
                    col = 0;


                    cells[row, ++col].Value = balanceGroup.GroupNameCleansed;
                    //Commitments
                    cells[row, ++col].Value = commitmentsRow.MonthToDateAmount;
                    cells[row, ++col].Value = commitmentsRow.YearToDateAmount;
                    //Contributions
                    cells[row, ++col].Value = contributionsRow.MonthToDateAmount;
                    cells[row, ++col].Value = contributionsRow.YearToDateAmount;
                    //Unfunded Commitments
                    cells[row, ++col].Value = unfundedCommitmentsRow.MonthToDateAmount;
                    cells[row, ++col].Value = unfundedCommitmentsRow.YearToDateAmount;
                    //Distributions
                    cells[row, ++col].Value = distributionsRow.MonthToDateAmount;
                    cells[row, ++col].Value = distributionsRow.YearToDateAmount;
                    //Total Ending Balance
                    cells[row, ++col].Value = totalEndingBalanceRow.MonthToDateAmount;
                    cells[row, ++col].Value = totalEndingBalanceRow.YearToDateAmount;
                }

                //format
                ws.Cells[3, 2, row, col].Style.Numberformat.Format = "_-$* #,##0.00_-;-$* #,##0.00_-;_-$* \"-\"??_-;_-@_-";

                //fit
                ws.Cells[ws.Dimension.Address].AutoFitColumns();

                excel.Save();
            }
        }

        public sealed class BalanceNames
        {
            public const string Commitments = "Commitments";
            public const string Contributions = "Contributions";
            public const string UnfundedCommitments = "Unfunded Commitments";
            public const string Distributions = "Distributions";

            public static readonly string[] Names = new[]
            {
                Commitments,
                Contributions,
                UnfundedCommitments,
                Distributions
            };
        }

        private class BalanceGroup
        {
            public string GroupName { get; set; } = "";
            public string GroupNameCleansed
            {
                get
                {
                    if (GroupName.Contains("#0"))
                    {
                        return GroupName.Replace("#0", "#");
                    }

                    return GroupName;
                }
            }
            public List<BalanceRecord> Records { get; set; } = new List<BalanceRecord>();
            public BalanceRecord Total { get; set; } = new BalanceRecord();

            public static BalanceGroup FromRecord(Record r)
            {
                var g = new BalanceGroup()
                {
                    GroupName = r.Name
                };

                return g;
            }
        }

        private class BalanceRecord
        {
            public string Name { get; set; } = "";
            public decimal? MonthToDateAmount { get; set; }
            public decimal? YearToDateAmount { get; set; }

            public static BalanceRecord FromRecord(Record r)
            {
                var record = new BalanceRecord()
                {
                    Name = r.Name,
                    MonthToDateAmount = r.MtdValue,
                    YearToDateAmount = r.YtdValue,
                };

                return record;
            }
        }

        public class Record
        {
            public string Name { get; set; } = "";
            public decimal? MtdValue { get; set; }
            public decimal? YtdValue { get; set; }

            public bool IsEmpty => string.IsNullOrWhiteSpace(Name)
                && MtdValue == null
                && YtdValue == null;

            public bool IsGroupHeader => !string.IsNullOrWhiteSpace(Name)
                && MtdValue == null
                && YtdValue == null;

            public bool IsBalanceRow => BalanceNames.Names.Contains(Name);

            public bool IsTotalRow => Name.StartsWith("Total ");

            public static Record FromExcel(ExcelRange cells, int row)
            {
                var r = new Record
                {
                    Name = cells[row, 1].GetValue<string>() ?? ""
                };

                var mtdValue = cells[row, 2].GetValue<string>();
                var ytdValue = cells[row, 3].GetValue<string>();

                if (decimal.TryParse(mtdValue, out var mtd))
                {
                    r.MtdValue = mtd;
                }

                if (decimal.TryParse(ytdValue, out var ytd))
                {
                    r.YtdValue = ytd;
                }

                return r;
            }
        }
    }
}
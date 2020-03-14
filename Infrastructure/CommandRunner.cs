using System;
using System.Threading.Tasks;
using ExcelConvert.Converters;
using Microsoft.Extensions.CommandLineUtils;
using Microsoft.Extensions.DependencyInjection;

namespace ExcelConvert.Infrastructure
{
    public class CommandRunner
    {
        private readonly IServiceProvider _serviceProvider;

        public CommandRunner(IServiceProvider serviceProvider)
        {
            _serviceProvider = serviceProvider;
        }

        public void Run(string[] args)
        {
            if (args == null || args.Length == 0)
            {
                args = new[] { "-h" };
            }

            //SEE: https://msdn.microsoft.com/en-us/magazine/mt763239.aspx
            var app = new CommandLineApplication(throwOnUnexpectedArg: true)
            {
                FullName = "\nExcel Converter",
                Description = "Excel Converter",
                Name = "ExcelConvert.exe",
                ExtendedHelpText = "\n\nExample:\n\n  ExcelConvert.exe --convert quick-book-balance-sheet --i \"C:\\test-sheet.xlsx\"\n"
            };

            var commandOptions = CommandOptions.ParseAndCreate(app);

            app.HelpOption("-? | -h | --help");

            //execute the commands
            app.OnExecute(async () =>
                {
                    return await ExecuteCommandsAsync(commandOptions);
                });

            app.Execute(args);
        }

        private async Task<int> ExecuteCommandsAsync(CommandOptions cmdOptions)
        {
            await ExecuteCommandConvert(cmdOptions);

            return 0;
        }


        private async Task ExecuteCommandConvert(CommandOptions commandOptions)
        {
            if (!commandOptions.Convert.HasValue())
            {
                return;
            }

            var optionName = nameof(commandOptions.Convert);
            var optionValue = commandOptions.Convert.Value();

            if (!Enum.TryParse(optionValue.Replace("-", ""), true, out CommandOptions.ConvertOption option))
            {
                throw new ArgumentOutOfRangeException(optionName, $"Unsupported {optionName} Option value - {optionValue}");
            }

            var runner = option switch
            {
                //SyncFrom
                CommandOptions.ConvertOption.QuickBookBalanceSheet => (IBatchService)_serviceProvider.GetService<QuickBookBalanceSheetConverter>(),
                //
                _ => throw new ArgumentOutOfRangeException(optionName, $"Un-handled {optionName} Option value: {optionValue}"),
            };

            await runner.RunAsync();
        }
    }
}

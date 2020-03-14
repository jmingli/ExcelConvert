using Microsoft.Extensions.CommandLineUtils;

namespace ExcelConvert.Infrastructure
{
    public class AppArgs
    {
        public string? InFileName { get; set; }
        public string? OutFileName { get; set; }

        public static AppArgs FromEntity(string[] args)
        {
            var app = new CommandLineApplication() { };
            var commandOptions = CommandOptions.ParseAndCreate(app);

            app.Execute(args);

            var appArgs = FromEntity(commandOptions);

            return appArgs;
        }

        public static AppArgs FromEntity(CommandOptions commandOptions)
        {
            var appArgs = new AppArgs();

            if (commandOptions.InFile.HasValue())
            {
                appArgs.InFileName = commandOptions.InFile.Value();
            }

            if (commandOptions.OutFile.HasValue())
            {
                appArgs.OutFileName = commandOptions.OutFile.Value();
            }

            return appArgs;
        }
    }
}

using ExcelConvert.Infrastructure;

namespace ExcelConvert
{
    internal class Program
    {
        private static void Main(string[] args)
        {
            var appArgs = AppArgs.FromEntity(args);

            var startup = new Startup(appArgs);

            var commandRunner = new CommandRunner(startup.ServiceProvider);

            commandRunner.Run(args);
        }
    }
}

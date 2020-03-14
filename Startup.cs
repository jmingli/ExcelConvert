using System.Linq;
using System.Reflection;
using ExcelConvert.Infrastructure;
using Microsoft.Extensions.DependencyInjection;

namespace ExcelConvert
{
    internal class Startup
    {
        private readonly AppArgs _appArgs;
        public ServiceProvider ServiceProvider { get; private set; }

        public Startup(AppArgs appArgs)
        {
            _appArgs = appArgs ?? new AppArgs();

            // Create a service collection and configure our dependencies
            var serviceCollection = new ServiceCollection();

            ConfigureServices(serviceCollection);

            ServiceProvider = serviceCollection.BuildServiceProvider();
        }

        private void ConfigureServices(IServiceCollection services)
        {
            services.AddSingleton<AppArgs>(_appArgs);

            var iBatchServiceType = typeof(IBatchService);
            var assembly = typeof(IBatchService).GetTypeInfo().Assembly;
            var batchServices = assembly.ExportedTypes
                .Where(x => x.IsClass)
                .Where(x => x.IsPublic)
                .Where(x => !x.IsAbstract)
                .AsEnumerable();

            foreach (var t in batchServices)
            {
                if (iBatchServiceType.IsAssignableFrom(t))
                {
                    services.AddTransient(t);
                }
            }
        }
    }
}

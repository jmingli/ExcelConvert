using System.Threading.Tasks;

namespace ExcelConvert.Infrastructure
{
    public interface IBatchService
    {
        Task RunAsync();
    }
}

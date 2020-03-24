using Microsoft.Extensions.Configuration;
using Microsoft.Extensions.DependencyInjection;

namespace UploadClient.Installers
{
    public class HealthChecksInstaller : IInstaller
    {
        public void InstallServices(IServiceCollection services, IConfiguration configuration)
        {
            services.AddHealthChecks();
               // .AddDbContextCheck<DataContext>()
              //  .AddCheck<RedisHealthCheck>("Redis");
        }
    }
}

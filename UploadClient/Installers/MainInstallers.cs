using Microsoft.Extensions.Configuration;
using Microsoft.Extensions.DependencyInjection;
using UploadClient.Models.Convertion;

namespace UploadClient.Installers
{
    public class MainInstallers : IInstaller
    {
        public void InstallServices(IServiceCollection services, IConfiguration configuration)
        {
            services.AddTransient<IConvertToExcel, ConvertToExcel>(); 

            services.AddControllersWithViews();

            services.AddSpaStaticFiles(configuration =>
            {
                configuration.RootPath = "ClientApp/build";
            });
        }
    }
}

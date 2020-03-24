using System;
using System.Collections.Generic;
using System.IO;
using System.Reflection;
using Microsoft.Extensions.Configuration;
using Microsoft.Extensions.DependencyInjection;
using Microsoft.OpenApi.Models;
using Swashbuckle.AspNetCore.Filters;

namespace UploadClient.Installers
{
    public class SwaggerInstaller : IInstaller
    {
        public void InstallServices(IServiceCollection services, IConfiguration configuration)
        {
            services.AddSwaggerGen(x =>
            {
                x.SwaggerDoc("v1", new OpenApiInfo { Title = "TB API", Version = "v1" });

                x.ExampleFilters();

                //var xmlFile = $"{Assembly.GetExecutingAssembly().GetName().Name}.xml";
                //var xmlPath = Path.Combine(AppContext.BaseDirectory, xmlFile);
                //x.IncludeXmlComments(xmlPath);
            });

            services.AddSwaggerExamplesFromAssemblyOf<Startup>();
        }
    }
}

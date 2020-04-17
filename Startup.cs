using System;
using System.Collections.Generic;
using System.Linq;
using System.Threading.Tasks;
using Microsoft.AspNetCore.Builder;
using Microsoft.AspNetCore.Hosting;
using Microsoft.AspNetCore.HttpsPolicy;
using Microsoft.AspNetCore.Mvc;
using Microsoft.Extensions.Configuration;
using Microsoft.Extensions.DependencyInjection;
using Microsoft.Extensions.Logging;
using Microsoft.Extensions.Options;
using Microsoft.Graph;
using ms_graph_app.Utils;

namespace ms_graph_app
{
  public class Startup
  {
    public Startup(IConfiguration configuration)
    {
      Configuration = configuration;
    }

    public IConfiguration Configuration { get; }

    // This method gets called by the runtime. Use this method to add services to the container.
    public void ConfigureServices(IServiceCollection services)
    {
      services.AddMvc().SetCompatibilityVersion(CompatibilityVersion.Version_2_1);
      var graphConfig = new GraphConfig();
      Configuration.Bind("GraphConfig", graphConfig);
      services.AddSingleton(graphConfig);
    }

    // This method gets called by the runtime. Use this method to configure the HTTP request pipeline.
    public void Configure(IApplicationBuilder app, IHostingEnvironment env, IApplicationLifetime lifetime)
    {
      lifetime.ApplicationStarted.Register(OnApplicationStarted);
      if(env.IsDevelopment())
      {
        app.UseDeveloperExceptionPage();
      }
      else
      {
        app.UseHsts();
        app.UseHttpsRedirection();
      }

      app.UseMvc();
    }

    public void OnApplicationStarted()
    {
      Console.WriteLine("started");

      var graphConfig = new GraphConfig();
      Configuration.Bind("GraphConfig", graphConfig);

      var graphHelper = new GraphHelper(graphConfig);
            _ = graphHelper.InitSubscription();
    }
  }
}

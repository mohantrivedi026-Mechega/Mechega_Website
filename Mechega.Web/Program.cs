using Microsoft.AspNetCore.Components.Web;
using Microsoft.AspNetCore.Components.WebAssembly.Hosting;
using Mechega.Web;

var builder = WebAssemblyHostBuilder.CreateDefault(args);
builder.RootComponents.Add<App>("#app");
builder.RootComponents.Add<HeadOutlet>("head::after");

// Use API origin for HttpClient in development to avoid same-origin issues
var baseAddr = builder.HostEnvironment.BaseAddress;
if (builder.HostEnvironment.IsDevelopment())
{
	// API runs on http://localhost:5000 during development
	baseAddr = "http://localhost:5000/";
}
builder.Services.AddScoped(sp => new HttpClient { BaseAddress = new Uri(baseAddr) });

await builder.Build().RunAsync();

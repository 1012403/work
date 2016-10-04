using Microsoft.Owin;
using Owin;

[assembly: OwinStartupAttribute(typeof(ZuelligPharma.Startup))]
namespace ZuelligPharma
{
    public partial class Startup
    {
        public void Configuration(IAppBuilder app)
        {
            ConfigureAuth(app);
        }
    }
}

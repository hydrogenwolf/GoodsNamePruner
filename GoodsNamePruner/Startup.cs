using Microsoft.Owin;
using Owin;

[assembly: OwinStartupAttribute(typeof(GoodsNamePruner.Startup))]
namespace GoodsNamePruner
{
    public partial class Startup
    {
        public void Configuration(IAppBuilder app)
        {
            ConfigureAuth(app);
        }
    }
}

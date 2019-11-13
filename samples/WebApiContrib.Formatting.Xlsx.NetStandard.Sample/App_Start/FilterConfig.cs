using System.Web.Mvc;

namespace WebApiContrib.Formatting.Xlsx.NetStandard.Sample
{
    public class FilterConfig
    {
        public static void RegisterGlobalFilters(GlobalFilterCollection filters)
        {
            filters.Add(new HandleErrorAttribute());
        }
    }
}

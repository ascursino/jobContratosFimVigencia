using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace jobContratosFimVigencia
{
    class mdConfig
    {
        //Comum
        public string user = "portalpmo";
        public string password = "b2w@123456";
        public string uri = "/_vti_bin/listdata.svc";
        public string linkContratoView = "/Lists/Contratos/DispForm.aspx";
        
        
        //Produção
        public string siteContratos = "http://wss.b2w/negocios_ti/adm_ti/contratos";
        public string domain = "lab2w";
        public string mailhost = "bwuolhub02.la.ad.b2w";
        public string mailfrom = "sharepointadmin@uoldiveo.com";

        //Desenv
        //public string siteContratos = "http://abcuniversity/contratos";
        //public string domain = "abcuniversity";
        //public string mailhost = "abcuniversity";
        //public string mailfrom = "sharepoint@abcuniversity.com";
        







    }
}

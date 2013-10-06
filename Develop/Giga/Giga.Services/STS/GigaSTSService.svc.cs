using System;
using System.Collections.Generic;
using System.IdentityModel;
using System.IdentityModel.Configuration;
using System.Linq;
using System.Runtime.Serialization;
using System.ServiceModel;
using System.Text;

namespace Giga.Services.STS
{
    // NOTE: You can use the "Rename" command on the "Refactor" menu to change the class name "GigaSTSService" in code, svc and config file together.
    // NOTE: In order to launch WCF Test Client for testing this service, please select GigaSTSService.svc or GigaSTSService.svc.cs at the Solution Explorer and start debugging.
    public class GigaSTSService : SecurityTokenService, IGigaSTSService
    {
        public GigaSTSService(SecurityTokenServiceConfiguration cfg) 
            : base(cfg)
        {
        }

        public void DoWork()
        {
        }

        protected override System.Security.Claims.ClaimsIdentity GetOutputClaimsIdentity(System.Security.Claims.ClaimsPrincipal principal, System.IdentityModel.Protocols.WSTrust.RequestSecurityToken request, Scope scope)
        {
            throw new NotImplementedException();
        }

        protected override Scope GetScope(System.Security.Claims.ClaimsPrincipal principal, System.IdentityModel.Protocols.WSTrust.RequestSecurityToken request)
        {
            throw new NotImplementedException();
        }
    }
}

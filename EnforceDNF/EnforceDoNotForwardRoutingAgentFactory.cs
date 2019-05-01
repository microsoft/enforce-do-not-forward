using Microsoft.Exchange.Data.Transport;
using Microsoft.Exchange.Data.Transport.Routing;

namespace EnforceDNF
{

    public class EnforceDoNotForwardRoutingAgentFactory : RoutingAgentFactory
    {
        public override RoutingAgent CreateAgent(SmtpServer server)
        {
            RoutingAgent enforceDoNotForwardAgent = new EnforceDoNotForwardRoutingAgent();
            return enforceDoNotForwardAgent;
        }
    }

}

using System;
using System.Net;
using System.Security.Authentication.ExtendedProtection;

namespace WebApiContrib.Formatting.Xlsx.Core.Tests.Fakes
{
    public class FakeTransport : TransportContext
    {
        public override ChannelBinding GetChannelBinding(ChannelBindingKind kind)
        {
            throw new NotImplementedException();
        }
    }
}

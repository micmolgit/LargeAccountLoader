using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

using Microsoft.Crm.Sdk.ServiceHelper;
using Microsoft.Xrm.Sdk;
using Microsoft.Xrm.Sdk.Client;

namespace LargeAccountLoader
{
    class DcrmConnector
    {
        private ServerConnection _serverConnection;
        private ServerConnection.Configuration _serverConfig;
        private OrganizationServiceProxy _serviceProxy;
        private ServiceContext _serviceContext;

        public ServiceContext SrvContext
        {
            get
            {
                if (_serviceContext == null)
                    _serviceContext = new ServiceContext(_serviceProxy);

                if (_serviceContext == null)
                {
                    throw new ArgumentNullException("ServiceContext", "ServiceContext Could not be created from OrganizationServiceProxy");
                }
                return _serviceContext;
            }
        }

        public OrganizationServiceProxy ServiceProxy
        {
            get
            {
                return _serviceProxy;
            }
        }

        public ServerConnection.Configuration ServiceConfig
        {
            get
            {
                return _serverConfig;
            }
        }

        public void Connect()
        {
            // Obtain the target organization's Web address and client logon
            // credentials from the user.
            _serverConnection = new ServerConnection();
            _serverConfig = _serverConnection.GetServerConfiguration();

            _serviceProxy = new OrganizationServiceProxy(_serverConfig.OrganizationUri, _serverConfig.HomeRealmUri, _serverConfig.Credentials, _serverConfig.DeviceCredentials);

            _serviceProxy.Authenticate();

            // This statement is required to enable early-bound type support.
            _serviceProxy.EnableProxyTypes();

            if (!_serviceProxy.IsAuthenticated)
            {
                throw new InvalidOperationException("Authentication could not be completed");
            }

            Console.WriteLine($"Successfully connected to : {_serverConfig.OrganizationUri}");
        }

        public void Disconnect()
        {
            _serviceProxy?.Dispose();
        }
    }
}

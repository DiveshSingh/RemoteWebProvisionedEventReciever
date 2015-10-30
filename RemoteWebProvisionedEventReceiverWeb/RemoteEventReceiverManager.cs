using Microsoft.SharePoint.Client;
using System;
using System.Collections.Generic;
using System.Linq;
using System.ServiceModel;
using System.ServiceModel.Channels;
using System.Web;

namespace RemoteWebProvisionedEventReceiverWeb
{
    public class RemoteEventReceiverManager
    {
        const string ReceiverName = "WebProvisionedEvent";
        ThemeHelper themeHelper;
        public RemoteEventReceiverManager()
        {
            themeHelper = new ThemeHelper();
        }
        /// <summary>
        /// deployes event receiver at site level
        /// </summary>
        /// <param name="clientContext"></param>
        public void AssociateWebProvisionedEventReceiver(ClientContext clientContext)
        {
            Web web = clientContext.Web;
            clientContext.Load(web);
            clientContext.Load(clientContext.Site);
            clientContext.Load(clientContext.Site.EventReceivers);
            clientContext.ExecuteQuery();
            EventReceiverDefinitionCreationInformation receiver = new EventReceiverDefinitionCreationInformation();
            receiver.EventType = EventReceiverType.WebProvisioned;
            OperationContext op = OperationContext.Current;
            Message msg = op.RequestContext.RequestMessage;
            receiver.ReceiverUrl = msg.Headers.To.ToString();
            receiver.ReceiverName = ReceiverName;
            receiver.Synchronization = EventReceiverSynchronization.Synchronous;
            receiver.SequenceNumber = 5000;
            //adding receiver the host web.
            // we are adding event receiver to site context so that event receriver fires for sub-sub-sites
            clientContext.Site.EventReceivers.Add(receiver);
            clientContext.ExecuteQuery();
            //upload theme related files
            themeHelper.DeployFiles(web);
            //delete oob composed looks
            themeHelper.DeleteOOBComposedLooks(clientContext);
            //apply theme and set default theme
            themeHelper.AddComposedLooksAndSetDefaultTheme(clientContext);


        }
        /// <summary>
        /// uninstalling webprovisioned event receiver during app unistallling
        /// </summary>
        /// <param name="clientContext"></param>
        public void RemoveWebProvisionedEventReceiver(ClientContext clientContext)
        {
            clientContext.Load(clientContext.Site);
            clientContext.Load(clientContext.Site.EventReceivers);
            clientContext.ExecuteQuery();
            var recievers = clientContext.Web.EventReceivers.FirstOrDefault(i => i.ReceiverName == ReceiverName);
            if(recievers!=null)
            {
                recievers.DeleteObject();
                clientContext.ExecuteQuery();
            }
        }
        /// <summary>
        /// uploading themes and apply defualt for new sub web just provisioned
        /// </summary>
        /// <param name="clientContext"></param>
        public void WebProvisionedEventReceiver(ClientContext clientContext)
        {
            themeHelper.DeleteOOBComposedLooks(clientContext);
            themeHelper.AddComposedLooksAndSetDefaultTheme(clientContext);
        }
    }
}
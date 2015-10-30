using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using Microsoft.SharePoint.Client;
using Microsoft.SharePoint.Client.EventReceivers;

namespace RemoteWebProvisionedEventReceiverWeb.Services
{
    public class AppEventReceiver : IRemoteEventService
    {   
        /// <summary>
        /// Handles app events that occur after the app is installed or upgraded, or when app is being uninstalled.
        /// </summary>
        /// <param name="properties">Holds information about the app event.</param>
        /// <returns>Holds information returned from the app event.</returns>
        public SPRemoteEventResult ProcessEvent(SPRemoteEventProperties properties)
        {
            SPRemoteEventResult result = new SPRemoteEventResult();
            switch (properties.EventType)
            {
               
                case SPRemoteEventType.WebProvisioned:
                    WebProvisionedMethod(properties);
                    break;
                case SPRemoteEventType.AppInstalled:
                    AppInstalledMethod(properties);
                    break;
                case SPRemoteEventType.AppUpgraded:
                    // not implemented
                    break;
                case SPRemoteEventType.AppUninstalling:
                    AppUnistallingMethod(properties);
                    break;
                case SPRemoteEventType.WebAdding:
                    // you can implement webaddding event
                    break;
                case SPRemoteEventType.WebDeleting:
                    //you can implemet web deleting event if needed.
                default:
                    break;
            }


            return result;
        }

        /// <summary>
        /// This method is a required placeholder, but is not used by app events.
        /// </summary>
        /// <param name="properties">Unused.</param>
        public void ProcessOneWayEvent(SPRemoteEventProperties properties)
        {
            throw new NotImplementedException();
        }
        /// <summary>
        /// Takes care of appinstalled method
        /// and installs web provisioned eventreceiver
        /// </summary>
        /// <param name="_properties"></param>
        private void AppInstalledMethod(SPRemoteEventProperties _properties)
        {
            using (ClientContext _clientContext = TokenHelper.CreateAppEventClientContext(_properties, false))
            {
                if(_clientContext!=null)
                {
                    new RemoteEventReceiverManager().AssociateWebProvisionedEventReceiver(_clientContext);
                }
            }
        }
        /// <summary>
        /// takes care of app uninstalling to remove web provisioned event receiver
        /// </summary>
        /// <param name="_properties"></param>
        private void AppUnistallingMethod(SPRemoteEventProperties _properties)
        {
            using (ClientContext _clientContext = TokenHelper.CreateAppEventClientContext(_properties, false))
            {
                if (_clientContext != null)
                {
                    new RemoteEventReceiverManager().RemoveWebProvisionedEventReceiver(_clientContext);
                }
            }
        }
        /// <summary>
        /// web provisioned method to upload themes and apply default theme
        /// delete OOb composed looks
        /// </summary>
        /// <param name="_properties"></param>
        private void WebProvisionedMethod(SPRemoteEventProperties _properties)
        {
            using (ClientContext _clientContext = TokenHelper.CreateRemoteEventReceiverClientContext(_properties))// creating remoteeventreceiver client context
            {
                if (_clientContext != null)
                {
                    new RemoteEventReceiverManager().WebProvisionedEventReceiver(_clientContext);
                }
            }
        }
    }

}

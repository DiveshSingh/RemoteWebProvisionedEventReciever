using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;
using Microsoft.SharePoint.Client;
using System.Web.Hosting;
using System.IO;

namespace RemoteWebProvisionedEventReceiverWeb
{
   
    public class ThemeHelper
    {
        List<ThemeProperties> themeProperties;
        /// <summary>
        /// This function is responsible to upload theme files which are in App
        /// to catalog list
        /// </summary>
        /// <param name="web"></param>
        public void DeployFiles(Web web)
        {
            try
            {
                List themesCatalog = web.GetCatalog(123);
                Folder rootFolder = themesCatalog.RootFolder;
                web.Context.Load(rootFolder);
                web.Context.Load(rootFolder.Folders);
                web.Context.ExecuteQuery();
                Folder folder15 = rootFolder.Folders.FirstOrDefault(f => f.Name == "15");
                #region GetFiles
                string path = HostingEnvironment.MapPath(string.Format("~/Resources"));
                // Get Theme Directories from the app physical path on IIS 
                string[] themeDirectories = Directory.GetDirectories(path);
                foreach (string themeDirUrl in themeDirectories)
                {
                    //am considering theme directory name as the theme name
                    // so getting theme directory name
                    string themeFolderName = themeDirUrl.Remove(0, path.Length + 1);
                    string[] themeFiles = Directory.GetFiles(themeDirUrl);
                    uploadFile(web, themeFiles, folder15, themeFolderName);
                }

                #endregion
            }
            catch (Exception ex)
            {
                // implment loggin
            }
        }
        /// <summary>
        /// uploading theme related files and capturing properites of theme
        /// </summary>
        /// <param name="web"></param>
        /// <param name="themeFiles"></param>
        /// <param name="folder15"></param>
        /// <param name="themeFolderName"></param>
        private void uploadFile(Web web, string[] themeFiles, Folder folder15, string themeFolderName)
        {
            ThemeProperties themeProp = new ThemeProperties();
            themeProp.ThemeName = themeFolderName;
            foreach (string file in themeFiles )
            {
                FileCreationInformation themeFile = new FileCreationInformation();
                themeFile.Content = System.IO.File.ReadAllBytes(file);
                themeFile.Url = folder15.ServerRelativeUrl + "/" + System.IO.Path.GetFileName(file);
                themeFile.Overwrite = true;
                Microsoft.SharePoint.Client.File uploadFile = folder15.Files.Add(themeFile);
                string fileExtn = Path.GetExtension(file);
                switch (fileExtn.ToLower())
                {
                    case ".spcolor":
                        themeProp.SPColorUrl = themeFile.Url;
                        break;
                    case ".spfont":
                        themeProp.FontUrl = themeFile.Url;
                        break;

                    case ".png":
                    case ".jpg":
                        themeProp.ImageUrl = themeFile.Url;
                        break;
                    default:
                        break;
                }
            }
            web.Context.ExecuteQuery();
            if(themeProperties!=null)
            {
                themeProperties = new List<ThemeProperties>();
            }
            // adding all theme file properties to the list to add these items to composed looks
            themeProperties.Add(themeProp);
        }
        /// <summary>
        /// deletes oob composed looks
        /// </summary>
        /// <param name="clientContext"></param>
        public void DeleteOOBComposedLooks(ClientContext clientContext)
        {
            try
            {
                List designCatalogList = clientContext.Web.GetCatalog(124);
                clientContext.Load(designCatalogList);
                clientContext.ExecuteQuery();
                ListItemCollection itemCollection = GetDesignCatalogItems(clientContext.Web,designCatalogList);
                foreach (ListItem item in itemCollection.ToList())
                {
                    item.DeleteObject();
                    designCatalogList.Update();
                }
                clientContext.ExecuteQuery();
            }
            catch (Exception ex)
            {
            }
            //throw new NotImplementedException();
        }
        /// <summary>
        /// 
        /// </summary>
        /// <param name="clientContext"></param>
        public void AddComposedLooksAndSetDefaultTheme(ClientContext clientContext)
        {
            List designCatalogList = clientContext.Web.GetCatalog(124);
            Web web = clientContext.Web;
            clientContext.Load(web);
            clientContext.ExecuteQuery();
            string masterPageUrl = string.Format("{0}/_catalogs/masterpage/oslo.master", web.ServerRelativeUrl);
            int displayOrder = 1;
            //if it is newly created sub web get the theme properites from the root web.
            if (themeProperties != null)
            {
                themeProperties = GetThemeListItems(clientContext);
            }

            // create items in composed looks
            foreach (ThemeProperties theme in themeProperties)
            {
                ListItemCreationInformation itemInfo = new ListItemCreationInformation();
                ListItem item = designCatalogList.AddItem(itemInfo);
                item["DisplayOrder"] = displayOrder;
                item["Name"] = theme.ThemeName;
                item["Title"] = theme.ThemeName;
                item["ThemeUrl"] = theme.SPColorUrl;
                item["FontSchemeUrl"] = theme.ImageUrl;
                item["ImageUrl"] = theme.ImageUrl;
                item["MasterPageUrl"] = masterPageUrl;
                displayOrder++;
                item.Update();

            }
            clientContext.ExecuteQuery();
            var defaultTheme = themeProperties.FirstOrDefault(i => i.ThemeName == "CustomDefaultTheme");
            if (defaultTheme != null)
            {
                ListItemCreationInformation itemInfo = new ListItemCreationInformation();
                ListItem item = designCatalogList.AddItem(itemInfo);
                item["DisplayOrder"] = 0;
                item["Name"] = "Current";
                item["Title"] = "Current";
                item["ThemeUrl"] = defaultTheme.SPColorUrl;
                item["FontSchemeUrl"] = defaultTheme.ImageUrl;
                item["ImageUrl"] = defaultTheme.ImageUrl;
                item["MasterPageUrl"] = masterPageUrl;
                item.Update();
                designCatalogList.Update();
                //applying the theme
                clientContext.Web.ApplyTheme(defaultTheme.SPColorUrl, defaultTheme.FontUrl, defaultTheme.ImageUrl, true);
                clientContext.ExecuteQuery();

            }
            // throw new NotImplementedException();
        }
        /// <summary>
        /// Get composed look list items
        /// </summary>
        /// <param name="clientContext"></param>
        /// <returns></returns>
        private List<ThemeProperties> GetThemeListItems(ClientContext clientContext)
        {
            List designCatalogList = clientContext.Site.RootWeb.GetCatalog(124);// Get composed looks from the roob web to create in subweb just got provisioned
            clientContext.Load(designCatalogList);
            clientContext.ExecuteQuery();
            ListItemCollection itemCollection = GetDesignCatalogItems(clientContext.Site.RootWeb, designCatalogList);
            foreach (ListItem item in itemCollection)
            {
                try
                {
                    if (item["Title"].ToString() != "Current")
                    {
                        themeProperties.Add(new ThemeProperties
                        {
                            ThemeName = item["Title"].ToString(),
                            ImageUrl = ((FieldUrlValue)item["ImageUrl"]).Description,//  description contains relative url 
                            SPColorUrl = ((FieldUrlValue)item["ThemeUrl"]).Description,
                            FontUrl = ((FieldUrlValue)item["FontSchemeUrl"]).Description
                        });
                    }
                }
                catch (Exception ex)
                {
                    // implement logging
                }

            }
            return themeProperties;
        }

        private ListItemCollection GetDesignCatalogItems(Web web,List designCatalogList)
        {
         
            CamlQuery query = new CamlQuery();
            query.ViewXml = "<View><OrderBy><FieldRef Name='Name'/></OrderBy></View>";
            ListItemCollection itemCollection = designCatalogList.GetItems(query);
            web.Context.Load(designCatalogList);
            web.Context.Load(itemCollection);
            web.Context.ExecuteQuery();
            return itemCollection;
        }
    }
}
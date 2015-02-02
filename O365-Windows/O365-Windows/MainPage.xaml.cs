//----------------------------------------------------------------------------------------------
//    Copyright 2014 Microsoft Corporation
//
//    Licensed under the Apache License, Version 2.0 (the "License");
//    you may not use this file except in compliance with the License.
//    You may obtain a copy of the License at
//
//      http://www.apache.org/licenses/LICENSE-2.0
//
//    Unless required by applicable law or agreed to in writing, software
//    distributed under the License is distributed on an "AS IS" BASIS,
//    WITHOUT WARRANTIES OR CONDITIONS OF ANY KIND, either express or implied.
//    See the License for the specific language governing permissions and
//    limitations under the License.
//----------------------------------------------------------------------------------------------

using Microsoft.IdentityModel.Clients.ActiveDirectory;
using Microsoft.Office365.Discovery;
using Microsoft.Office365.SharePoint.CoreServices;
using Microsoft.Office365.SharePoint.FileServices;
using O365_Windows.Helpers;
using O365_Windows.Model;
using System;
using System.Collections.Generic;
using System.Collections.ObjectModel;
using System.IO;
using System.Linq;
using System.Runtime.InteropServices.WindowsRuntime;
using Windows.Foundation;
using Windows.Foundation.Collections;
using Windows.UI.Popups;
using Windows.UI.Xaml;
using Windows.UI.Xaml.Controls;
using Windows.UI.Xaml.Controls.Primitives;
using Windows.UI.Xaml.Data;
using Windows.UI.Xaml.Input;
using Windows.UI.Xaml.Media;
using Windows.UI.Xaml.Navigation;

// The Blank Page item template is documented at http://go.microsoft.com/fwlink/?LinkId=234238

namespace O365_Windows
{
    /// <summary>
    /// An empty page that can be used on its own or navigated to within a Frame.
    /// </summary>
    public sealed partial class MainPage : Page
    {   
        public ObservableCollection<MyFile> Files { get; set; }
        
        public MainPage()
        {
            Files = new ObservableCollection<MyFile>();

            this.InitializeComponent();            
        }

        private async void btnGetMyFiles_Click(object sender, RoutedEventArgs e)
        {
            Files.Clear();

            DiscoveryClient discoveryClient = new DiscoveryClient(
                    async ()=>
                    {
                        var authResult = await AuthenticationHelper.GetAccessToken(AuthenticationHelper.DiscoveryServiceResourceId);

                        return authResult.AccessToken;
                    }
                );

            var appCapabilities = await discoveryClient.DiscoverCapabilitiesAsync();

            var myFilesCapability = appCapabilities
                                    .Where(s => s.Key == "MyFiles")
                                    .Select(p=>new {Key=p.Key, ServiceResourceId=p.Value.ServiceResourceId, ServiceEndPointUri=p.Value.ServiceEndpointUri})
                                    .FirstOrDefault();
                                    
            
            if(myFilesCapability != null)
            {
                SharePointClient myFilesClient = new SharePointClient(myFilesCapability.ServiceEndPointUri,
                    async()=>
                    {
                        var authResult = await AuthenticationHelper.GetAccessToken(myFilesCapability.ServiceResourceId);

                        return authResult.AccessToken;
                    });

                var myFilesResult = await myFilesClient.Files.ExecuteAsync();

                do
                {
                    var myFiles = myFilesResult.CurrentPage;
                    foreach (var myFile in myFiles)
                    {
                        Files.Add(new MyFile { Name = myFile.Name });
                    }

                    myFilesResult = await myFilesResult.GetNextPageAsync();

                } while (myFilesResult != null);

                if(Files.Count == 0)
                {
                    Files.Add(new MyFile { Name = "No files to display!" });
                }

                

            }
            else
            {
                MessageDialog dialog = new MessageDialog(string.Format("This Windows app does not have access to users' files. Please contact your administrator."));
                await dialog.ShowAsync();
            }
           
        }
    }
}

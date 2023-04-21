using McTools.Xrm.Connection;
using Microsoft.Crm.Sdk.Messages;
using Microsoft.Xrm.Sdk;
using Microsoft.Xrm.Sdk.Query;
using Microsoft.Xrm.Tooling.Connector;
using Newtonsoft.Json;
using Newtonsoft.Json.Linq;
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Text.RegularExpressions;
using System.Threading.Tasks;
using System.Windows.Forms;
using XrmToolBox.Extensibility;

namespace Amend_Connection_Reference
{
    public partial class MyPluginControl : PluginControlBase
    {
        private Settings mySettings;
        private CrmServiceClient _crmServiceClient;
        private ConnectionDetail _connectionDetail;
        public MyPluginControl()
        {
            InitializeComponent();
            
        }

        private void MyPluginControl_Load(object sender, EventArgs e)
        {
            ShowInfoNotification("This is a notification that can lead to XrmToolBox repository", new Uri("https://github.com/MscrmTools/XrmToolBox"));

            // Loads or creates the settings for the plugin
            if (!SettingsManager.Instance.TryLoad(GetType(), out mySettings))
            {
                mySettings = new Settings();

                LogWarning("Settings not found => a new settings file has been created!");
            }
            else
            {
                LogInfo("Settings found and loaded");
            }
        }

        private void tsbClose_Click(object sender, EventArgs e)
        {
            CloseTool();
        }
        /// <summary>
        /// This event occurs when the plugin is closed
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void MyPluginControl_OnCloseTool(object sender, EventArgs e)
        {
            // Before leaving, save the settings
            SettingsManager.Instance.Save(GetType(), mySettings);
        }

        /// <summary>
        /// This event occurs when the connection has been updated in XrmToolBox
        /// </summary>
        public override void UpdateConnection(IOrganizationService newService, ConnectionDetail detail, string actionName, object parameter)
        {
            base.UpdateConnection(newService, detail, actionName, parameter);
            _connectionDetail = detail;
            if (mySettings != null && detail != null)
            {
                mySettings.LastUsedOrganizationWebappUrl = detail.WebApplicationUrl;
                LogInfo("Connection has changed to: {0}", detail.WebApplicationUrl);
            }
        }

        private void toolStripButton1_Click(object sender, EventArgs e)
        {
            try
            {
                GetCloudFlows();
                GetConnectionReference();
            }
            catch (Exception ex)
            {
                MessageBox.Show($"Error: {ex.Message}.");
            }

        }

        private void GetCloudFlows()
        {
            WorkAsync(new WorkAsyncInfo
            {
                Message = "Getting Cloud Flows and Connection Reference",
                Work = (worker, args) =>
                {
                    var fetchData = new
                    {
                        category = "5",
                        statecode = "1",
                        ismanaged = "1"
                    };
                    var fetchXml = $@"
                    <fetch>
                      <entity name='workflow'>
                        <attribute name='statecode' />
                        <attribute name='name' />
                        <attribute name='category' />
                        <attribute name='clientdata' />
                        <attribute name='workflowid' />
                        <order attribute='name' />
                        <filter>
                          <condition attribute='category' operator='eq' value='{fetchData.category/*5*/}'/>
                          <condition attribute='statecode' operator='eq' value='{fetchData.statecode/*1*/}'/>
                        </filter>
                      </entity>
                    </fetch>";
                    args.Result = Service.RetrieveMultiple(new FetchExpression(fetchXml));
                },
                PostWorkCallBack = (args) =>
                {
                    if (args.Error != null)
                    {
                        MessageBox.Show(args.Error.ToString(), "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                    }
                    var result = args.Result as EntityCollection;
                    if (result != null)
                    {
                        WorkflowDetails workflowDetails;
                        foreach (var item in result.Entities)
                        {
                            workflowDetails = new WorkflowDetails();
                            workflowDetails.clientData = item.Attributes["clientdata"].ToString();
                            workflowDetails.workflowName = item.Attributes["name"].ToString();
                            workflowDetails.workflowid = Guid.Parse(item.Attributes["workflowid"].ToString());
                            comboBox1.Items.Add(workflowDetails);
                        }
                    }
                }
            });
        }

        private void GetConnectionReference()
        {
            WorkAsync(new WorkAsyncInfo
            {
                Message = "Getting Cloud Flows and Connection Reference",
                Work = (worker, args) =>
                {
                    var fetchData = new
                    {
                        connectionreferencelogicalname = "%sharedcommondataserviceforapps%"
                    };
                    var fetchXml = $@"
                        <fetch>
                          <entity name='connectionreference'>
                            <attribute name='connectionreferencedisplayname' />
                            <attribute name='connectionreferencelogicalname' />
                            <attribute name='ownerid' />
                            <order attribute='connectionreferencedisplayname' />
                            <filter>
                              <condition attribute='connectionreferencelogicalname' operator='like' value='{fetchData.connectionreferencelogicalname/*%sharedcommondataserviceforapps%*/}' />
                            </filter>
                          </entity>
                        </fetch>";
                    args.Result = Service.RetrieveMultiple(new FetchExpression(fetchXml));
                },
                PostWorkCallBack = (args) =>
                {
                    if (args.Error != null)
                    {
                        MessageBox.Show(args.Error.ToString(), "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                    }
                    var result = args.Result as EntityCollection;
                    if (result != null)
                    {
                        ConnectionReferenceName connRefName;
                        EntityReference owner;
                        foreach (var item in result.Entities)
                        {
                            connRefName = new ConnectionReferenceName();
                            connRefName.connectionRefLogicalName = item.Attributes["connectionreferencelogicalname"].ToString();
                            connRefName.connectionRefDisplayName = item.Attributes["connectionreferencedisplayname"].ToString();
                            owner = (EntityReference)item.Attributes["ownerid"];
                            connRefName.ownerId = owner.Id.ToString();
                            comboBox2.Items.Add(connRefName);
                        }
                        MessageBox.Show("Retrieved Cloud Flows and Connection References.");
                    }
                }
            });
        }

        private void comboBox1_Click(object sender, EventArgs e)
        {
            //GetCloudFlows();
        }

        private void comboBox2_Click(object sender, EventArgs e)
        {
            //GetConnectionReference();
        }

        private void button1_Click(object sender, EventArgs e)
        {
            var flowDetails = (WorkflowDetails)this.comboBox1.SelectedItem;
            var connectionRef = (ConnectionReferenceName)this.comboBox2.SelectedItem;
            var selectIndexConnectionReference = this.comboBox2.SelectedIndex;
            var selectIndexFlow = this.comboBox1.SelectedIndex;
            if (selectIndexConnectionReference != -1 && selectIndexFlow != -1)
            {
                this.UpdateConnectionReference(flowDetails, connectionRef);
            }
            else
            {
                MessageBox.Show("Please populate Cloud Flow and Connection Reference post clicking on Load Cloud Flows and Connection References Ribbon Button.");
            }
        }
        private void UpdateConnectionReference(WorkflowDetails flowDetails, ConnectionReferenceName connectionRefDetails)
        {
            WorkAsync(new WorkAsyncInfo
            {
                Message = "Updating Connection Reference",
                Work = (worker, args) =>
                {
                    //var connectionDetail = new ConnectionDetail();
                    var service = _connectionDetail.GetCrmServiceClient(true);

                    service.CallerId = Guid.Parse(connectionRefDetails.ownerId);

                    string clientData = flowDetails.clientData;
                    var connectionRef = "";

                    var connectionReferences = new ConnectionReferences();
                    var sharedCommondataserviceforapps = new SharedCommondataserviceforapps();
                    var connect = new Connection();
                    var api = new Api();

                    api.name = "shared_commondataserviceforapps";
                    connect.connectionReferenceLogicalName = connectionRefDetails.connectionRefLogicalName;
                    sharedCommondataserviceforapps.runtimeSource = "embedded";
                    sharedCommondataserviceforapps.connection = connect;
                    sharedCommondataserviceforapps.api = api;
                    connectionReferences.shared_commondataserviceforapps = sharedCommondataserviceforapps;
                    connectionRef = JsonConvert.SerializeObject(connectionReferences);
                    var newConnectionReference = JObject.Parse(connectionRef);

                    string pattern = "shared_commondataserviceforapps_[0-9]+";
                    Regex rg = new Regex(pattern);
                    var cloudflowClientData = JObject.Parse(clientData);

                    // Updating exisiting connection reference with New connection reference.
                    cloudflowClientData["properties"]["connectionReferences"] = newConnectionReference;

                    // Replacing all shared_commondataserviceforapps_[0-9] with shared_commondataserviceforapps.
                    clientData = Regex.Replace(cloudflowClientData.ToString(), pattern, "shared_commondataserviceforapps");

                    var workflowToUpdate = new Entity("workflow");
                    workflowToUpdate["workflowid"] = flowDetails.workflowid;
                    workflowToUpdate["clientdata"] = clientData;

                    // Updating the workflow.
                    service.Update(workflowToUpdate);
                    MessageBox.Show($"Connection Reference for Cloud Flow {flowDetails.workflowName} has been updated to {connectionRefDetails.connectionRefDisplayName}.");
                },
                PostWorkCallBack = (args) =>
                {
                    if (args.Error != null)
                    {
                        MessageBox.Show(args.Error.ToString(), "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                    }
                    var result = args.Result as EntityCollection;
                    if (result != null)
                    {
                        MessageBox.Show($"Connection Reference has been updated.");
                    }
                }
            });
        }

        
    }
    public class Connection
    {
        public string connectionReferenceLogicalName { get; set; }
    }

    public class ConnectionReferenceName
    {
        public string connectionRefLogicalName { get; set; }
        public string connectionRefDisplayName { get; set; }
        public string ownerId { get; set; }
    }

    public class WorkflowDetails
    {
        public string workflowName { get; set; }
        public string clientData { get; set; }
        public Guid workflowid { get; set; }
    }

    public class Api
    {
        public string name { get; set; }
    }

    public class SharedCommondataserviceforapps
    {
        public string runtimeSource { get; set; }
        public Connection connection { get; set; }
        public Api api { get; set; }
    }

    public class ConnectionReferences
    {
        public SharedCommondataserviceforapps shared_commondataserviceforapps { get; set; }
    }
}
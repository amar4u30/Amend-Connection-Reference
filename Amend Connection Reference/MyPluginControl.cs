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
                reloadFlowsandConnectionReferences();
            }
            catch (Exception ex)
            {
                MessageBox.Show($"Error: {ex.Message}.");
            }

        }
        private void reloadFlowsandConnectionReferences()
        {
            this.comboBox2.Items.Clear();
            this.comboBox1.Items.Clear();
            this.checkedListBox1.Items.Clear();
            GetCloudFlows();
            GetConnectionReference();
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
                            checkedListBox1.Items.Add(workflowDetails);
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
                        connectorid = "/providers/Microsoft.PowerApps/apis/shared_commondataserviceforapps"
                    };
                    var fetchXml = $@"
                        <fetch>
                          <entity name='connectionreference'>
                            <attribute name='connectionreferencedisplayname' />
                            <attribute name='connectionreferencelogicalname' />
                            <attribute name='ownerid' />
                            <order attribute='connectionreferencedisplayname' />
                            
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
                            connRefName.connectionRefDisplayName = $"{item.Attributes["connectionreferencedisplayname"]}, Logical Name: {item.Attributes["connectionreferencelogicalname"]}";
                            owner = (EntityReference)item.Attributes["ownerid"];
                            connRefName.ownerId = owner.Id.ToString();
                            comboBox2.Items.Add(connRefName);
                        }
                        //MessageBox.Show("Retrieved Cloud Flows and Connection References.");
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
            var selectIndexFlow = this.checkedListBox1.SelectedIndex;
            if (selectIndexConnectionReference != -1 && selectIndexFlow != -1)
            {
                var checkedItems = this.checkedListBox1.CheckedItems;
                this.UpdateConnectionReference(checkedItems, connectionRef);
                //this.PopulateDataGridView();
            }
            else
            {
                MessageBox.Show("Please Select Cloud Flow and Connection Reference.");
            }
        }
        private void UpdateConnectionReference(CheckedListBox.CheckedItemCollection checkedItems, ConnectionReferenceName connectionRefDetails)
        {
            WorkAsync(new WorkAsyncInfo
            {
                Message = "Updating Connection Reference",
                Work = (worker, args) =>
                {
                    //var connectionDetail = new ConnectionDetail();
                    var service = _connectionDetail.GetCrmServiceClient(true);
                    var updateFlag = false;
                    var systemCreatedConnectionReference = false;
                    service.CallerId = Guid.Parse(connectionRefDetails.ownerId);
                    try
                    {
                        if (false/*connectionRefDetails.connectionRefLogicalName.Contains("commondataserviceforapps")*/)
                        {
                            foreach (WorkflowDetails item in checkedItems)
                            {
                                string clientData = item.clientData;
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
                                workflowToUpdate["workflowid"] = item.workflowid;
                                workflowToUpdate["clientdata"] = clientData;

                                // Updating the workflow.
                                service.Update(workflowToUpdate);
                            }
                        }
                        else
                        {
                            var connectionRef = "";
                            var existingStart = connectionRefDetails.connectionRefLogicalName.IndexOf("_");
                            var existingEnd = connectionRefDetails.connectionRefLogicalName.IndexOf("_", existingStart + 1);
                            if(existingEnd !=-1 && existingEnd != -1)
                            {
                                systemCreatedConnectionReference = true;
                                var existingConnectionName = connectionRefDetails.connectionRefLogicalName.Substring(existingStart + 1, existingEnd - existingStart - 1);
                                foreach (WorkflowDetails item in checkedItems)
                                {
                                    string clientData = item.clientData;

                                    var connectioName = existingConnectionName.Replace("shared", null);
                                    string pattern = $"shared_{connectioName}_[0-9]+";
                                    clientData = Regex.Replace(clientData.ToString(), pattern, $"shared_{connectioName}");


                                    var indexOfString = clientData.IndexOf(existingConnectionName);
                                    if (indexOfString != -1)
                                    {
                                        var start = clientData.LastIndexOf("\"", indexOfString);
                                        var end = clientData.IndexOf("_", indexOfString);
                                        var connectionRefToBeReplaced = clientData.Substring(start + 1, end - start - 1);

                                        pattern = $"{connectionRefToBeReplaced}_[a-zA-Z0-9]+";
                                        clientData = Regex.Replace(clientData.ToString(), pattern, connectionRefDetails.connectionRefLogicalName);

                                        var workflowToUpdate = new Entity("workflow");
                                        workflowToUpdate["workflowid"] = item.workflowid;
                                        workflowToUpdate["clientdata"] = clientData;

                                        // Updating the workflow.
                                        service.Update(workflowToUpdate);
                                        updateFlag = true;
                                    }
                                }
                            }
                        }
                        if (!systemCreatedConnectionReference)
                        {
                            MessageBox.Show($"Opps! Looks like you are trying to update with a manually created Connection Reference. \nThe tool is not configured to handle manually created Connection Reference, please use a Connection Reference created from a Cloud Flow.");
                        }
                        else if (updateFlag)
                        {
                            MessageBox.Show($"The Selected Cloud Flow(s) having matching Connection Reference with {connectionRefDetails.connectionRefDisplayName.Split(',')[0]} have been updated.");
                        }
                        else
                        {
                            MessageBox.Show($"Sorry! The Selected Cloud Flow(s) doesnt have a matching Connection Reference with {connectionRefDetails.connectionRefDisplayName.Split(',')[0]} hence couldnt be updated.");
                        }
                    }
                    catch (Exception exc)
                    {
                        MessageBox.Show($"Unexpected Error Occured, Message: {exc.Message}");
                    }
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
                    this.checkedListBox1.Items.Clear();
                    this.GetCloudFlows();
                }
            });
        }

        private void checkBox1_CheckedChanged(object sender, EventArgs e)
        {
            for (int i = 0; i < this.checkedListBox1.Items.Count; i++)
            {
                checkedListBox1.SetItemChecked(i, checkBox1.Checked);
            }
        }

        private void tableLayoutPanel1_Paint(object sender, PaintEventArgs e)
        {

        }

        private void checkedListBox1_SelectedIndexChanged(object sender, EventArgs e)
        {
            //PopulateDataGridView();
        }
        private void PopulateDataGridView()
        {
            this.dataGridView1.Rows.Clear();
            var checkedItems = checkedListBox1.CheckedItems;
            var connectionList = new List<ConnectionObject>();
            if (checkedItems.Count > 0)
            {
                WorkAsync(new WorkAsyncInfo
                {
                    Message = "Fetching Connection Reference(s)",
                    Work = (worker, args) =>
                    {
                        foreach (WorkflowDetails item in checkedItems)
                        {
                            var firstTime = true;
                            var cloudflowClientData = JObject.Parse(item.clientData);
                            var filterString = new StringBuilder("");
                            var connectionReferencesList = (JObject)cloudflowClientData["properties"]?["connectionReferences"];
                            foreach (KeyValuePair<string, JToken> connectionReference in connectionReferencesList)
                            {
                                string keyName = connectionReference.Key;
                                var value = cloudflowClientData["properties"]?["connectionReferences"]?[keyName]?["connection"]?["connectionReferenceLogicalName"]?.ToString();
                                filterString.Append($"<condition attribute='connectionreferencelogicalname' operator='eq' value='{value}' />");
                            }

                            var connectionReferenceDisplayName = "";
                            var fetchXml = $@"
                                        <fetch>
                                          <entity name='connectionreference'>
                                            <attribute name='connectionreferencedisplayname' />
                                            <attribute name='connectionreferencelogicalname' />
                                            <attribute name='ownerid' />
                                            <order attribute='connectionreferencedisplayname' />
                                            <filter type='or'>
                                             {filterString}
                                            </filter>
                                          </entity>
                                        </fetch>";
                            EntityCollection connectionReferences = Service.RetrieveMultiple(new FetchExpression(fetchXml));
                            foreach (var connectionReference in connectionReferences.Entities)
                            {
                                connectionReferenceDisplayName = connectionReference.Attributes["connectionreferencedisplayname"].ToString();
                                var connectionObj = new ConnectionObject()
                                {
                                    connectionReferenceLogicalName = connectionReferenceDisplayName,
                                    flowName = firstTime ? item.workflowName : ""
                                };
                                connectionList.Add(connectionObj);
                                firstTime = false;
                            }
                            
                        }
                    },
                    PostWorkCallBack = (args) =>
                    {
                        if (args.Error != null)
                        {
                            MessageBox.Show(args.Error.ToString(), "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                        }
                        foreach (var item in connectionList)
                        {
                            dataGridView1.Rows.Add(item.flowName, item.connectionReferenceLogicalName);
                        }
                    }
                });
                
            }
            else
            {
                MessageBox.Show("Please Select Cloud Flow(s).");
            }
        }

        private void button2_Click(object sender, EventArgs e)
        {
            PopulateDataGridView();
        }
    }
    public class Property
    {
        public List<string> connectionReferences { get; set; }
    }
    public class ConnectionObject
    {
        public string connectionReferenceLogicalName { get; set; }
        public string flowName { get; set; }
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
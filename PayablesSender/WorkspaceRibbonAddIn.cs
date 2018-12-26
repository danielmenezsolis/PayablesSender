using System;
using System.AddIn;
using System.Collections.Generic;
using System.IO;
using System.Net;
using System.Net.Security;
using System.Security.Cryptography.X509Certificates;
using System.ServiceModel;
using System.ServiceModel.Channels;
using System.Text;
using System.Windows.Forms;
using System.Xml;
using System.Xml.Linq;
using System.Xml.Serialization;
using PayablesSender.SOAPICCS;
using RightNow.AddIns.AddInViews;
using RightNow.AddIns.Common;


namespace PayablesSender
{
    public class WorkspaceRibbonAddIn : Panel, IWorkspaceRibbonButton
    {
        private IGlobalContext GlobalContext { get; set; }
        private IIncident IncidentRecord { get; set; }
        private IRecordContext RecordContext { get; set; }
        private RightNowSyncPortClient clientORN { get; set; }
        private int IncidentID { get; set; }
        private DATA_DS_ITEMSUP suppliers { get; set; }

        public WorkspaceRibbonAddIn(bool inDesignMode, IRecordContext RecordContext, IGlobalContext globalContext)
        {
            GlobalContext = globalContext;
            this.RecordContext = RecordContext;
        }
        public new void Click()
        {
            try
            {
                getSuppliers();
                Init();
                IncidentRecord = (IIncident)RecordContext.GetWorkspaceRecord(WorkspaceRecordType.Incident);
                IncidentID = IncidentRecord.ID;
                Cursor.Current = Cursors.WaitCursor;
                SendPayables();
                Cursor.Current = Cursors.Default;
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message + "Det" + ex.StackTrace);
            }

        }
        public void SendPayables()
        {
            try
            {
                ClientInfoHeader clientInfoHeader = new ClientInfoHeader();
                APIAccessRequestHeader aPIAccessRequest = new APIAccessRequestHeader();
                clientInfoHeader.AppID = "Query Example";
                String queryString = "SELECT Airport,ItemNumber,Incident.CreatedTime,Itinerary.ETA,Itinerary.ETATime,Itinerary.ETD,Itinerary.ETDTime,Itinerary.ATA,Itinerary.ATATime,Itinerary.ATD,Itinerary.ATDTime,Incident.Customfields.c.sr_type.LookupName,Paquete,Incident.CustomFields.co.Aircraft.AircraftType1.ICAODesignator,Services.ItemNumber,Precio,Incident.LookupName,fuel_id,IdProveedor,ID,Incident.StatusWithType.Status.Name,IVA FROM CO.Services WHERE Incident =" + IncidentID;
                clientORN.QueryCSV(clientInfoHeader, aPIAccessRequest, queryString, 10000, "|", false, false, out CSVTableSet queryCSV, out byte[] FileData);
                foreach (CSVTable table in queryCSV.CSVTables)
                {
                    String[] rowData = table.Rows;
                    foreach (String data in rowData)
                    {
                        Char delimiter = '|';
                        string[] substrings = data.Split(delimiter);
                        string Airport = substrings[0];
                        string ItemNumber = substrings[1];
                        string IncidentCreatedTime = substrings[2];
                        string ItineraryETA = substrings[3];
                        string ItineraryETATime = substrings[4];
                        string ItineraryETD = substrings[5];
                        string ItineraryETDTime = substrings[6];
                        string ItineraryATA = substrings[7];
                        string ItineraryATATime = substrings[8];
                        string ItineraryATD = substrings[9];
                        string ItineraryATDTime = substrings[10];
                        string SRType = substrings[11];
                        string Paquete = substrings[12] == "1" ? "Y" : "N";
                        string ICAODesignator = substrings[13];
                        string ServicesItemNumber = substrings[14];
                        string Precio = substrings[15];
                        string IncidentLookupName = substrings[16];
                        string fuel_id = substrings[17];
                        string IdProveedor = substrings[18];
                        string IDService = substrings[19];
                        string Status = substrings[20];
                        string IVA = substrings[21];

                        string envelope = "<soapenv:Envelope" +
                        "   xmlns:soapenv=\"http://schemas.xmlsoap.org/soap/envelope/\"" +
                        "   xmlns:mod=\"http://modelo.iccs/\">" +
                        "<soapenv:Header/>" +
                        "<soapenv:Body>" +
                        "<mod:createSR>" +
                           "<payload>" +
                               "<estructuraList>" +
                                "<AEROPUERTO>" + Airport + "</AEROPUERTO>" +
                                "<ATA_DATE>" + ItineraryATA + "</ATA_DATE>" +
                                "<ATA_TIME>" + ItineraryATATime + "</ATA_TIME>" +
                                "<ATD_DATE>" + ItineraryATD + "</ATD_DATE>" +
                                "<ATD_TIME>" + ItineraryATDTime + "</ATD_TIME>" +
                                "<CANTIDAD>1</CANTIDAD>" +
                                "<CODIGO_ITEM>" + ItemNumber + "</CODIGO_ITEM>" +
                                "<CODIGO_PAQUETE>" + ItemNumber + "</CODIGO_PAQUETE>" +
                                "<CUENTA_CONTABLE>N/A</CUENTA_CONTABLE>" +
                                "<ESTATUS>A</ESTATUS>" +
                                "<ETA_DATE>" + ItineraryETA + "</ETA_DATE>" +
                                "<ETA_TIME>" + ItineraryETATime + "</ETA_TIME>" +
                                "<ETD_DATE>" + ItineraryETD + "</ETD_DATE>" +
                                "<ETD_TIME>" + ItineraryETDTime + "</ETD_TIME>" +
                                "<FECHA>" + IncidentCreatedTime + "</FECHA>" +
                                "<IMPUESTO>" + IVA + "</IMPUESTO>" +
                                "<IS_PAQUETE>" + Paquete + "</IS_PAQUETE>" +
                                "<ITEM>" + ItemNumber + "</ITEM>" +
                                "<MATRICULA>" + ICAODesignator + "</MATRICULA>" +
                                "<PAQUETE>" + ItemNumber + "</PAQUETE>" +
                                "<PRECIO>" + Precio + "</PRECIO>" +
                                "<PROVEEDOR>" + IdProveedor + "</PROVEEDOR>" +
                                "<SR>" + IncidentLookupName + "</SR>" +
                                "<TIPO_COMBUSTIBLE></TIPO_COMBUSTIBLE>" +
                                "<TOTAL>" + Precio + "</TOTAL>" +
                                "<UOM>SER</UOM>" +
                                "<VOUCHER>" + GetVoucher(fuel_id) + "</VOUCHER>" +
                                "<idSR>" + IncidentID + "</idSR>";
                        envelope += getPayablesChild(IDService);

                        envelope += "</estructuraList>" +
                                    "</payload>" +
                                "</mod:createSR>" +
                            "</soapenv:Body>" +
                        "</soapenv:Envelope>";
                        byte[] byteArray = Encoding.UTF8.GetBytes(envelope);
                        GlobalContext.LogMessage(envelope);
                        // Construct the base 64 encoded string used as credentials for the service call
                        byte[] toEncodeAsBytes = System.Text.ASCIIEncoding.ASCII.GetBytes("itotal" + ":" + "Oracle123");
                        string credentials = System.Convert.ToBase64String(toEncodeAsBytes);
                        HttpWebRequest request =
                         (HttpWebRequest)WebRequest.Create("https://129.150.72.35:7002/ICCS/InterfacePort");
                        request.Method = "POST";
                        request.ContentType = "text/xml;charset=UTF-8";
                        request.ContentLength = byteArray.Length;
                        ServicePointManager.ServerCertificateValidationCallback = new RemoteCertificateValidationCallback(AcceptAllCertifications);
                        Stream dataStream = request.GetRequestStream();
                        dataStream.Write(byteArray, 0, byteArray.Length);
                        dataStream.Close();

                        // Write the xml payload to the request
                        XDocument doc;
                        XmlDocument docu = new XmlDocument();
                        string result;
                        // Get the response and process it; In this example, we simply print out the response XDocument doc;
                        using (WebResponse response = request.GetResponse())
                        {

                            using (Stream stream = response.GetResponseStream())
                            {
                                doc = XDocument.Load(stream);
                                result = doc.ToString();
                                XmlDocument xmlDoc = new XmlDocument();
                                xmlDoc.LoadXml(result);
                                XmlNamespaceManager nms = new XmlNamespaceManager(xmlDoc.NameTable);
                                nms.AddNamespace("ns0", "http://modelo.iccs/");
                                XmlNode desiredNode = xmlDoc.SelectSingleNode("//idResult", nms);
                                if (desiredNode != null)
                                {
                                    if (desiredNode.InnerText != "0")
                                    {
                                        XmlNode error = xmlDoc.SelectSingleNode("//resultDescription", nms);
                                        {
                                            if (error != null)
                                            {
                                                MessageBox.Show(error.InnerText);
                                            }
                                        }
                                    }
                                    else
                                    {
                                        MessageBox.Show("Created");
                                    }


                                }

                            }
                        }
                    }
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message + "Det" + ex.StackTrace);
            }
        }
        static bool AcceptAllCertifications(object sender, X509Certificate cert, X509Chain chain, SslPolicyErrors errors)
        {
            return true;
        }
        public bool Init()
        {
            try
            {
                bool result = false;
                EndpointAddress endPointAddr = new EndpointAddress(GlobalContext.GetInterfaceServiceUrl(ConnectServiceType.Soap));
                // Minimum required
                BasicHttpBinding binding = new BasicHttpBinding(BasicHttpSecurityMode.TransportWithMessageCredential);
                binding.Security.Message.ClientCredentialType = BasicHttpMessageCredentialType.UserName;
                binding.ReceiveTimeout = new TimeSpan(0, 10, 0);
                binding.MaxReceivedMessageSize = 1048576; //1MB
                binding.SendTimeout = new TimeSpan(0, 10, 0);
                // Create client proxy class
                clientORN = new RightNowSyncPortClient(binding, endPointAddr);
                // Ask the client to not send the timestamp
                BindingElementCollection elements = clientORN.Endpoint.Binding.CreateBindingElements();
                elements.Find<SecurityBindingElement>().IncludeTimestamp = false;
                clientORN.Endpoint.Binding = new CustomBinding(elements);
                // Ask the Add-In framework the handle the session logic
                GlobalContext.PrepareConnectSession(clientORN.ChannelFactory);
                if (clientORN != null)
                {
                    result = true;
                }

                return result;
            }
            catch (Exception ex)
            {
                MessageBox.Show("Error en INIT: " + ex.Message);
                return false;

            }
        }
        public string getPayablesChild(string Service)
        {
            try
            {
                string hijoslist = "";
                ClientInfoHeader clientInfoHeader = new ClientInfoHeader();
                APIAccessRequestHeader aPIAccessRequest = new APIAccessRequestHeader();
                clientInfoHeader.AppID = "Query Example";
                String queryString = "SELECT Services.Airport,ItemNumber,Services.Incident.CreatedTime,Services.Itinerary.ETA,Services.Itinerary.ETATime,Services.Itinerary.ETD,Services.Itinerary.ETDTime,Services.Itinerary.ATA,Services.Itinerary.ATATime,Services.Itinerary.ATD,Services.Itinerary.ATDTime,Services.Incident.Customfields.c.sr_type.LookupName,Paquete,Services.Incident.CustomFields.co.Aircraft.AircraftType1.ICAODesignator,Services.ItemNumber,TicketAmount,Services.Incident.LookupName,Services.fuel_id,Supplier,ID,Services.Incident.StatusWithType.Status.Name,UOM,Services,Quantity,TicketAmount*Quantity FROM CO.Payables WHERE Services =" + Service;
                clientORN.QueryCSV(clientInfoHeader, aPIAccessRequest, queryString, 10000, "|", false, false, out CSVTableSet queryCSV, out byte[] FileData);
                foreach (CSVTable table in queryCSV.CSVTables)
                {
                    String[] rowData = table.Rows;
                    foreach (String data in rowData)
                    {
                        Char delimiter = '|';
                        string[] substrings = data.Split(delimiter);
                        string ServicesAirport = substrings[0];
                        string ItemNumber = substrings[1];
                        string ServicesIncidentCreatedTime = substrings[2];
                        string ServicesItineraryETA = substrings[3];
                        string ServicesItineraryETATime = substrings[4];
                        string ServicesItineraryETD = substrings[5];
                        string ServicesItineraryETDTime = substrings[6];
                        string ServicesItineraryATA = substrings[7];
                        string ServicesItineraryATATime = substrings[8];
                        string ServicesItineraryATD = substrings[9];
                        string ServicesItineraryATDTime = substrings[10];
                        string ServicesIncidentCustomfieldscSRType = substrings[11];
                        string Paquete = substrings[12] == "1" ? "Y" : "N"; ;
                        string ServicesIncidentICAODesignator = substrings[13];
                        string ServicesItemNumber = substrings[14];
                        string TicketAmount = substrings[15];
                        string ServicesIncidentLookupName = substrings[16];
                        string Servicesfuel_id = substrings[17];
                        string Supplier = substrings[18];
                        string ID = substrings[19];
                        string ServicesStatus = substrings[20];
                        string UOM = substrings[21];
                        string Services = substrings[22];
                        string Qty = string.IsNullOrEmpty(substrings[23]) ? "0" : substrings[23];
                        string Tot = string.IsNullOrEmpty(substrings[24]) ? "0" : substrings[24];
                        hijoslist += "<hijosList>" +
                                                   "<AEROPUERTO>" + ServicesAirport + "</AEROPUERTO>" +
                                                       "<ATA_DATE>" + ServicesItineraryATA + "</ATA_DATE>" +
                                                       "<ATA_TIME>" + ServicesItineraryATATime + "</ATA_TIME>" +
                                                       "<ATD_DATE>" + ServicesItineraryATD + "</ATD_DATE>" +
                                                       "<ATD_TIME>" + ServicesItineraryATDTime + "</ATD_TIME>" +
                                                       "<CANTIDAD>" + Qty + "</CANTIDAD>" +
                                                       "<CODIGO_ITEM>" + ItemNumber + "</CODIGO_ITEM>" +
                                                       "<CODIGO_PAQUETE>" + ServicesItemNumber + "</CODIGO_PAQUETE>" +
                                                       "<CUENTA_CONTABLE>N/A</CUENTA_CONTABLE>" +
                                                       "<ESTATUS>A</ESTATUS>" +
                                                       "<ETA_DATE>" + ServicesItineraryETA + "</ETA_DATE>" +
                                                       "<ETA_TIME>" + ServicesItineraryETATime + "</ETA_TIME>" +
                                                       "<ETD_DATE>" + ServicesItineraryETD + "</ETD_DATE>" +
                                                       "<ETD_TIME>" + ServicesItineraryETDTime + "</ETD_TIME>" +
                                                       "<FECHA>" + ServicesIncidentCreatedTime + "</FECHA>" +
                                                       "<IMPUESTO></IMPUESTO>" +
                                                       "<IS_PAQUETE>" + Paquete + "</IS_PAQUETE>" +
                                                       "<ITEM>" + ItemNumber + "</ITEM>" +
                                                       "<MATRICULA>" + ServicesIncidentICAODesignator + "</MATRICULA>" +
                                                       "<PAQUETE>" + ServicesItemNumber + "</PAQUETE>" +
                                                       "<PRECIO>" + TicketAmount + "</PRECIO>" +
                                                       "<PROVEEDOR>" + Supplier + "</PROVEEDOR>" +
                                                       "<SR>" + ServicesIncidentLookupName + "</SR>" +
                                                       "<TIPO_COMBUSTIBLE></TIPO_COMBUSTIBLE>" +
                                                       "<TOTAL>" + Tot + "</TOTAL>" +
                                                       "<UOM>" + UOM + "</UOM>" +
                                                       "<VOUCHER>" + GetVoucher(Servicesfuel_id) + "</VOUCHER>" +
                                                       "<idSR>" + IncidentID + "</idSR>" +
                                            "</hijosList>";
                    }
                }
                return hijoslist;
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message + "Det" + ex.StackTrace);
                return "";
            }
        }
        private string GetVoucher(string FuelId)
        {
            try
            {
                FuelId = string.IsNullOrEmpty(FuelId) ? "0" : FuelId;
                string voucher = "";
                //Liters * 3.7854 
                ClientInfoHeader clientInfoHeader = new ClientInfoHeader();
                APIAccessRequestHeader aPIAccessRequest = new APIAccessRequestHeader();
                clientInfoHeader.AppID = "Query Example";
                String queryString = "SELECT VoucherNumber FROM CO.Fueling WHERE ID = " + FuelId;
                clientORN.QueryCSV(clientInfoHeader, aPIAccessRequest, queryString, 1, "|", false, false, out CSVTableSet queryCSV, out byte[] FileData);
                foreach (CSVTable table in queryCSV.CSVTables)
                {
                    String[] rowData = table.Rows;
                    foreach (String data in rowData)
                    {
                        voucher = data;
                    }
                }
                return voucher;
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message + "Det" + ex.StackTrace);
                return "";
            }
        }
        private void getSuppliers()
        {
            try
            {

                string envelope = "<soap:Envelope " +
               "	xmlns:soap=\"http://www.w3.org/2003/05/soap-envelope\"" +
    "	xmlns:pub=\"http://xmlns.oracle.com/oxp/service/PublicReportService\">" +
     "<soap:Header/>" +
    "	<soap:Body>" +
    "		<pub:runReport>" +
    "			<pub:reportRequest>" +
    "			<pub:attributeFormat>xml</pub:attributeFormat>" +
    "				<pub:attributeLocale></pub:attributeLocale>" +
    "				<pub:attributeTemplate></pub:attributeTemplate>" +
    "				<pub:reportAbsolutePath>Custom/Integracion/XX_ITEM_SUPPLIER_ORG_REP.xdo</pub:reportAbsolutePath>" +
    "				<pub:sizeOfDataChunkDownload>-1</pub:sizeOfDataChunkDownload>" +
    "			</pub:reportRequest>" +
    "		</pub:runReport>" +
    "	</soap:Body>" +
    "</soap:Envelope>";
                byte[] byteArray = Encoding.UTF8.GetBytes(envelope);
                // Construct the base 64 encoded string used as credentials for the service call
                byte[] toEncodeAsBytes = ASCIIEncoding.ASCII.GetBytes("itotal" + ":" + "Oracle123");
                string credentials = Convert.ToBase64String(toEncodeAsBytes);
                // Create HttpWebRequest connection to the service
                HttpWebRequest request =
                 (HttpWebRequest)WebRequest.Create("https://egqy-test.fa.us6.oraclecloud.com:443/xmlpserver/services/ExternalReportWSSService");
                // Configure the request content type to be xml, HTTP method to be POST, and set the content length
                request.Method = "POST";

                request.ContentType = "application/soap+xml; charset=UTF-8;action=\"\"";
                request.ContentLength = byteArray.Length;
                // Configure the request to use basic authentication, with base64 encoded user name and password, to invoke the service.
                request.Headers.Add("Authorization", "Basic " + credentials);

                Stream dataStream = request.GetRequestStream();
                dataStream.Write(byteArray, 0, byteArray.Length);
                dataStream.Close();
                // Write the xml payload to the request
                XDocument doc;
                XmlDocument docu = new XmlDocument();
                string result;
                using (WebResponse response = request.GetResponse())
                {
                    using (Stream stream = response.GetResponseStream())
                    {
                        doc = XDocument.Load(stream);
                        result = doc.ToString();
                        XmlDocument xmlDoc = new XmlDocument();
                        xmlDoc.LoadXml(result);

                        XmlNamespaceManager nms = new XmlNamespaceManager(xmlDoc.NameTable);
                        nms.AddNamespace("env", "http://schemas.xmlsoap.org/soap/envelope/");
                        nms.AddNamespace("ns2", "http://xmlns.oracle.com/oxp/service/PublicReportService");

                        XmlNode desiredNode = xmlDoc.SelectSingleNode("//ns2:runReportReturn", nms);
                        if (desiredNode != null)
                        {
                            if (desiredNode.HasChildNodes)
                            {
                                for (int i = 0; i < desiredNode.ChildNodes.Count; i++)
                                {
                                    if (desiredNode.ChildNodes[i].LocalName == "reportBytes")
                                    {
                                        byte[] data = Convert.FromBase64String(desiredNode.ChildNodes[i].InnerText);
                                        string decodedString = Encoding.UTF8.GetString(data);
                                        XmlTextReader reader = new XmlTextReader(new System.IO.StringReader(decodedString));
                                        reader.Read();
                                        XmlSerializer serializer = new XmlSerializer(typeof(DATA_DS_ITEMSUP));
                                        suppliers = (DATA_DS_ITEMSUP)serializer.Deserialize(reader);
                                    }
                                }
                            }
                        }
                    }
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show("GetSupliers" + ex.Message + "Detalle: " + ex.StackTrace);
            }
        }
        private string getIdSupplier(string airport, string name)
        {
            string Supplier = "No Supplier";

            var lista = suppliers.G_N_ITEMSUP.Find(x => (x.ORGANIZATION_CODE.Trim().ToUpper() == airport.ToUpper()));
            if (lista != null)
            {
                foreach (G_1_ITEMSUP item in lista.G_1_ITEMSUP)
                {
                    if (item.PARTY_NAME.ToUpper() == name.ToUpper())
                    {
                        Supplier = item.PARTY_NUMBER;
                    }
                }
            }

            return Supplier;
        }

    }

    [AddIn("Send Payables", Version = "1.0.0.0")]
    public class WorkspaceRibbonButtonFactory : IWorkspaceRibbonButtonFactory
    {
        public IGlobalContext globalContext { get; set; }

        public IWorkspaceRibbonButton CreateControl(bool inDesignMode, IRecordContext RecordContext)
        {
            return new WorkspaceRibbonAddIn(inDesignMode, RecordContext, globalContext);
        }

        public System.Drawing.Image Image32
        {
            get { return Properties.Resources.jigsaw32; }
        }


        public System.Drawing.Image Image16
        {
            get { return Properties.Resources.jigsaw16; }
        }


        public string Text
        {
            get { return "Send Payables"; }
        }


        public string Tooltip
        {
            get { return "Send Payables"; }
        }

        public bool Initialize(IGlobalContext GlobalContext)
        {
            globalContext = GlobalContext;
            return true;
        }


    }


    [XmlRoot(ElementName = "G_1_ITEMSUP")]
    public class G_1_ITEMSUP
    {
        [XmlElement(ElementName = "ASL_ID")]
        public string ASL_ID { get; set; }
        [XmlElement(ElementName = "VENDOR_ID")]
        public string VENDOR_ID { get; set; }
        [XmlElement(ElementName = "PARTY_ID")]
        public string PARTY_ID { get; set; }
        [XmlElement(ElementName = "PARTY_NUMBER")]
        public string PARTY_NUMBER { get; set; }
        [XmlElement(ElementName = "PARTY_NAME")]
        public string PARTY_NAME { get; set; }
        [XmlElement(ElementName = "INVENTORY_ITEM_ID")]
        public string INVENTORY_ITEM_ID { get; set; }
        [XmlElement(ElementName = "ITEM_NUMBER")]
        public string ITEM_NUMBER { get; set; }
        [XmlElement(ElementName = "DESCRIPTION")]
        public string DESCRIPTION { get; set; }
        [XmlElement(ElementName = "PRIMARY_UOM_CODE")]
        public string PRIMARY_UOM_CODE { get; set; }
        [XmlElement(ElementName = "VENDOR_SITE_CODE")]
        public string VENDOR_SITE_CODE { get; set; }
    }

    [XmlRoot(ElementName = "G_N_ITEMSUP")]
    public class G_N_ITEMSUP
    {
        [XmlElement(ElementName = "ORGANIZATION_CODE")]
        public string ORGANIZATION_CODE { get; set; }
        [XmlElement(ElementName = "G_1_ITEMSUP")]
        public List<G_1_ITEMSUP> G_1_ITEMSUP { get; set; }
    }

    [XmlRoot(ElementName = "DATA_DS_ITEMSUP")]
    public class DATA_DS_ITEMSUP
    {
        [XmlElement(ElementName = "G_N_ITEMSUP")]
        public List<G_N_ITEMSUP> G_N_ITEMSUP { get; set; }
    }



}
using System;
using System.Collections.Generic;
using System.Linq;
using System.Xml.Linq;
using System.Threading;
using System.Windows.Forms;
using System.Data;

using Intense.Common;
using Intense.Common.Enums;


namespace UPS_Integration
{
    public class XmlExport : Intense.Interfaces.IInternalAction
    {
        public void Configure(Dictionary<string, object> parameters)
        {
            throw new NotImplementedException("ExtensionNoConfiguration"); //(Resources.ExtensionNoConfiguration);
        }

        public bool Execute(Dictionary<string, object> parameters)
        {
            #region Stwórz Logger lub otwórz istniejący

            int logID = Intense.Common.ErrorLogger.GetThreadLogID();

            if (logID < 0)
                logID = Intense.Common.ErrorLogger.CreateNewLogger();
            #endregion

            parameters.Add("LogID", (object)logID);

            try
            {
                Intense.Common.ErrorLogger.Logger.Init(logID, Intense.Common.Enums.LogProcessType.Other, true, string.Format("UPS Exporting XML"), 0);
                if (!parameters.ContainsKey("CURRENTID"))
                    throw new Exception("No CurrentID");
           /*     for(int i = 0; i < parameters.Count; i++)
                {
                    ErrorLogger.Logger.LogEvent(string.Format("TEST: {0} - {1}", parameters.ElementAt(i).Key, parameters.ElementAt(i).Value), (ErrorPriority)1, string.Empty);
                }

                ErrorLogger.Logger.LogEvent(string.Format("TEST - Current id: {0}", parameters["CURRENTID"]), (ErrorPriority)1, string.Empty);
                */                

                if (Context.Instance().AgentMode)
                {
                    this.AsyncXmlExport((object)parameters);
                }
                else
                {
                    Thread thread = new Thread(new ParameterizedThreadStart(this.AsyncXmlExport));
                    thread.CurrentCulture = Thread.CurrentThread.CurrentCulture;
                    thread.CurrentUICulture = Thread.CurrentThread.CurrentUICulture;
                    thread.Name = "ImportPositions";
                    thread.Start((object)parameters);
                    while (thread.IsAlive)
                    {
                        Application.DoEvents();
                        Thread.Sleep(100);
                    }
                }
                return true;
            }
            catch (Exception ex)
            {
                ErrorLogger.Logger.LogEvent(ex, (ErrorPriority)3, string.Empty);
                return false;
            }
            finally
            {
                ErrorLogger.Logger.Close(logID, (LogProcessType)1, true);
            }
        }


        private void AsyncXmlExport(object arg)
        {
            Dictionary<string, object> parameters = arg as Dictionary<string, object>;

            try
            {
                ErrorLogger.SetThreadLogID(Convert.ToInt32(parameters["LogID"]));


                //ErrorLogger.Logger.LogEvent(string.Format("Downloading data"), (ErrorPriority)1, string.Empty);
                ErrorLogger.Logger.OpenLevel("Downloading data", false);

                #region Downloading data

                int? currentId = Convert.ToInt32(parameters["CURRENTID"]);
                if (!currentId.HasValue)
                    throw new Exception("No ID of Shipment document");// (string.Format(Properties.Resources.ColumnNotContainsValue, "PRODUCTCODE"));
                if(currentId == 0)
                {
                    int? taskId = Convert.ToInt32(parameters["CURRENTTASKID"]);
                    string sql2 = string.Format(@"SELECT Tsk_DocID FROM Tasks WHERE Tsk_ID = {0}", taskId);
                    currentId = Convert.ToInt32(Intense.Common.SQL.GetSQLQueryResult(sql2));
                }

                if (currentId == 0)
                    throw new Exception("CurrentID = 0");

                ErrorLogger.Logger.LogEvent(string.Format("Current id: {0}", currentId), (ErrorPriority)1, string.Empty);

                //MessageBox.Show("Current ID: " + currentId ); //TEST

                string sql = string.Format(@"
SELECT * FROM IG_FIXIT_UPSDataCollection ({0})
    ", currentId);

                DataTable sqlResult = Intense.Common.SQL.GetSQLQuery(sql);

                if(sqlResult == null)
                    ErrorLogger.Logger.LogEvent(string.Format("SQL result amount: NULL"), (ErrorPriority)1, string.Empty);
                else
                    ErrorLogger.Logger.LogEvent(string.Format("SQL result amount: {0}", sqlResult.Rows.Count), (ErrorPriority)1, string.Empty);
                //MessageBox.Show("SQL result amount: " + sqlResult.Rows.Count); //TEST

            
                if (sqlResult == null || sqlResult.Rows.Count == 0)
                    throw new Exception("No data in table");
                
                #endregion

                
                                
                foreach (DataRow row in sqlResult.Rows)
                {
                    if (Intense.Common.ErrorLogger.Logger.IsCancelled)
                        throw new Exception("Cancelled");// (string.Format(Properties.Resources.Stop, "INVENTORYNO"));
                    #region Reading data
                    ErrorLogger.Logger.OpenLevel("Reading data", false);

                    int? courierId = row["CourierID"] as int?;  //31 - UPS; 75 - UPS Access Point
                    int? shipmentType = row["ShipmentType"] as int?;
                    string filePath = row["FilePath"] as string;
                    string waybillNumber = row["WaybillNumber"] as string;
                    string shipmentOption = row["ShipmentOption"] as string;

                    // Ship From
                    string shipFromCompany = row["ShipFromCompany"] as string;
                    string shipFromAttention = row["ShipFromAttention"] as string;
                    string shipFromAddress1 = row["ShipFromAddress1"] as string;
                    string shipFromCountryTerritory = row["ShipFromCountryTerritory"] as string;
                    string shipFromPostalCode = row["ShipFromPostalCode"] as string;
                    string shipFromCity = row["ShipFromCity"] as string;
                    string shipFromPhone = row["ShipFromPhone"] as string;
                    string shipFromEmail = row["ShipFromEmail"] as string;

                    // Ship To
                    string shipToCompany = row["ShipToCompany"] as string;
                    string shipToAttention = row["ShipToAttention"] as string;
                    string shipToAddress1 = row["ShipToAddress1"] as string;
                    string shipToCountryTerritory = row["ShipToCountryTerritory"] as string;
                    string shipToPostalCode = row["ShipToPostalCode"] as string;
                    string shipToCity = row["ShipToCity"] as string;
                    string shipToPhone = row["ShipToPhone"] as string;
                    string shipToEmail = row["ShipToEmail"] as string;

                    //AccessPoint
                    string pointCompany = row["PointCompany"] as string;
                    string pointAddress1 = row["PointAddress1"] as string;
                    string pointCountryTerritory = row["PointCountryTerritory"] as string;
                    string pointPostalCode = row["PointPostalCode"] as string;
                    string pointCity = row["PointCity"] as string;

                    //Shipment Information
                    string shipperNumber = row["ShipperNumber"] as string;
                    string serviceType = row["ServiceType"] as string;
                    string packageType = row["PackageType"] as string;
                    int? packageAmount = row["PackageAmount"] as int?;
                    int? shipmentActualWeight = row["ShipmentActualWeight"] as int?;
                    string descriptionOfGoods = row["DescriptionOfGoods"] as string;
                    string reference = row["Reference"] as string;
                    //COD
                    int? codIndicator = row["CODIndicator"] as int?;
                    decimal? codAmount = row["CODAmount"] as decimal?;
                    string codCurrency = row["CODCurrency"] as string;
                    string codBillingOption = row["BillingOption"] as string;
                    string codInvoice = row["CODInvoice"] as string;
                    //QVNOption
                    string NotifEmail = row["NotifEmail"] as string;
                    string PickupCompany = row["PickupCompany"] as string;
                    string PickupEmail = row["PickupEmail"] as string;

                    string apNotifEmail = row["APNotifEmail"] as string;
                    string apNotifLanguage = row["APNotifLanguage"] as string;

                    

                    ErrorLogger.Logger.CloseLevel();
                    #endregion

                    #region Exporting XML

                    if (courierId != 31 && courierId != 75)
                    {
                        MessageBox.Show("INCORRECT COURIER!!!!");
                        throw new Exception("Incorret courier. ");
                    }
                    else
                    {
                        ErrorLogger.Logger.OpenLevel("Exporting XML", false);
                        //XNamespace empNM = "http://www.ups.com/XMLSchema/CT/WorldShip/ImpExp/ShipmentImport/v1_0_0";
                        XNamespace empNM = "x-schema: \\\\192.168.102.2\\_storage_\\Logistyka\\Import_ups_xml\\Openshipments.xdr";
                        XDocument xmlFile = new XDocument(
                            new XDeclaration("1.0", "WINDOWS-1250", null)
                            );

                        //, new XElement("OpenShipments")


                        XElement openShipments = new XElement(empNM + "OpenShipments");
                        xmlFile.Add(openShipments);

                        XElement openShipment = new XElement(empNM + "OpenShipment", new XAttribute("ProcessStatus", ""), new XAttribute("ShipmentOption", "EU"));
                        openShipments.Add(openShipment);

                        XElement shipTo = new XElement(empNM + "ShipTo"
                           , new XElement(empNM + "CompanyOrName", shipToCompany)
                           , new XElement(empNM + "Attention", shipToAttention)
                           , new XElement(empNM + "Address1", shipToAddress1)
                           , new XElement(empNM + "CountryTerritory", shipToCountryTerritory)
                           , new XElement(empNM + "PostalCode", shipToPostalCode)
                           , new XElement(empNM + "CityOrTown", shipToCity)
                           , new XElement(empNM + "Telephone", shipToPhone)
                           , new XElement(empNM + "EmailAddress", shipToEmail)
                           );
                        openShipment.Add(shipTo);

                        XElement shipFrom = new XElement(empNM + "ShipFrom"
                           , new XElement(empNM + "CompanyOrName", shipFromCompany)
                           , new XElement(empNM + "Attention", shipFromAttention)
                           , new XElement(empNM + "Address1", shipFromAddress1)
                           , new XElement(empNM + "CountryTerritory", shipFromCountryTerritory)
                           , new XElement(empNM + "PostalCode", shipFromPostalCode)
                           , new XElement(empNM + "CityOrTown", shipFromCity)
                           , new XElement(empNM + "Telephone", shipFromPhone)
                           //     , new XElement(empNM + "Email", shipFromEmail)   // to check
                           );
                        openShipment.Add(shipFrom);

                        #region Dodanie adresu AccessPoint
                        if (shipmentType != null && shipmentType != 13698 && shipmentType != 13699 && courierId == 75)
                            if (!string.IsNullOrEmpty(pointCompany))
                            {
                                XElement accessPoint = new XElement(empNM + "AccessPoint"
                                    , new XElement(empNM + "CompanyOrName", pointCompany)
                                    , new XElement(empNM + "Address1", pointAddress1)
                                    , new XElement(empNM + "CountryTerritory", pointCountryTerritory)
                                    , new XElement(empNM + "PostalCode", pointPostalCode)
                                    , new XElement(empNM + "CityOrTown", pointCity)
                                    );
                                openShipment.Add(accessPoint);
                            }
                        #endregion

                        XElement shipmentInformation = new XElement(empNM + "ShipmentInformation"
                            , new XElement(empNM + "ShipperNumber", shipperNumber)
                            , new XElement(empNM + "DescriptionOfGoods", reference)
                            , new XElement(empNM + "ServiceType", "ST")
                            );
                        openShipment.Add(shipmentInformation);

                        #region Add Return service for UPS - AccessPoint Customer -> Fixit transport

                        if (courierId == 75 && shipmentType == 13698)
                        {
                            XElement returnService = new XElement(empNM + "ReturnService"
                                , new XElement(empNM + "Options", "PRL")
                                , new XElement(empNM + "MerchandiseDescOfPackage", reference)
                            );
                            shipmentInformation.Add(returnService);
                        }else if(courierId == 31 && shipmentType == 13698)
                        {
                            XElement returnService = new XElement(empNM + "ReturnService"
                               , new XElement(empNM + "Options", "RS1")
                               , new XElement(empNM + "MerchandiseDescOfPackage", reference)
                           );
                            shipmentInformation.Add(returnService);
                        }
                        #endregion

                        #region Notification otpions
                        string customerEmail, fixitEmail;
                        if (shipmentType != null && shipmentType != 13698 && shipmentType != 13699)
                        {
                            customerEmail = shipToEmail;
                            fixitEmail = shipFromEmail;
                        }
                        else
                        {
                            customerEmail = shipFromEmail;
                            fixitEmail = shipToEmail;
                        }

                        XElement qVNOption = new XElement(empNM + "QVNOption"
                        , new XElement(empNM + "QVNRecipientAndNotificationTypes"
                            , new XElement(empNM + "EMailAddress", customerEmail)
                            , new XElement(empNM + "Ship", "Y")
                            , new XElement(empNM + "Exception", "Y")
                            , new XElement(empNM + "Delivery", "N")
                        )
                        , new XElement(empNM + "QVNRecipientAndNotificationTypes"
                            , new XElement(empNM + "EMailAddress", fixitEmail)
                            , new XElement(empNM + "Ship", "N")
                            , new XElement(empNM + "Exception", "Y")
                            , new XElement(empNM + "Delivery", "Y")
                        )
                        , new XElement(empNM + "ShipFromCompanyOrName", "UPS - FIXIT SA")
                        , new XElement(empNM + "FailedEMailAddress", "logisytka@fixit.pl")
                        );
                        shipmentInformation.Add(qVNOption);
                        #endregion

                        XElement billingOption = new XElement(empNM + "BillingOption", "PP");
                        shipmentInformation.Add(billingOption);

                        if (!string.IsNullOrEmpty(pointCompany))
                        {
                            XElement holdAtUPSAccessPointOption = new XElement(empNM + "HoldatUPSAccessPointOption"
                                , new XElement(empNM + "APNotificationType", "1")
                                , new XElement(empNM + "APNotificationValue", customerEmail)
                                , new XElement(empNM + "APNotificationFailedEmailAddress", customerEmail)
                                , new XElement(empNM + "APNotificationLanguage", "ENG")
                                );
                            shipmentInformation.Add(holdAtUPSAccessPointOption);
                        }

                        XElement package = new XElement(empNM + "Package"
                            , new XElement(empNM + "PackageType", "CP")
                            , new XElement(empNM + "Weight", shipmentActualWeight)
                            , new XElement(empNM + "TrackingNumber", waybillNumber)
                            );
                        if (codIndicator == 1)
                        {
                            XElement reference2 = new XElement(empNM + "Reference1", codInvoice);
                            XElement cod = new XElement(empNM + "COD"
                                    , new XElement(empNM + "Amount", codAmount)
                                    , new XElement(empNM + "Currency", codCurrency)

                                    );
                            package.Add(reference2);
                            package.Add(cod);
                        }
                        openShipment.Add(package);



                        xmlFile.Save(filePath + waybillNumber + ".xml");
                        //MessageBox.Show("Filepath: " + filePath + waybillNumber + ".xml"); //TEST
                    }
                    ErrorLogger.Logger.CloseLevel();
                    #endregion

                }
            }
            catch(Exception ex)
            {
                ErrorLogger.Logger.LogEvent(ex, (ErrorPriority)3, string.Empty);
                
            }
            finally
            {
                ErrorLogger.Logger.CloseLevel();
            }

        }
    }
}
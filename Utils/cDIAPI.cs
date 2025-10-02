using CrystalDecisions.CrystalReports.Engine;
using NLog;
using SAPbobsCOM;
using System;
using System.Collections.Generic;
using System.Data;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace NortConsultingTasks.Utils
{
    public class cDIAPI
    {

        public static SAPbobsCOM.Company ConnectionCompany(string companyDB = "")
        {
            string serverIP = Properties.Settings.Default.ServerSAP;
            string SLDServer = Properties.Settings.Default.SLDServer;
            string BBDD = Properties.Settings.Default.BBDD;
            string userBBDD = Properties.Settings.Default.UserBBDD;
            string PassBBDD = EncryptionHelper.DecryptPass(Properties.Settings.Default.PwdBBDD);
            string userSAP = Properties.Settings.Default.UserSAP;
            string PassSAP = EncryptionHelper.DecryptPass(Properties.Settings.Default.PwdSAP);

            SAPbobsCOM.Company oCompany = new SAPbobsCOM.Company();
            oCompany.Server = serverIP;
            oCompany.SLDServer = SLDServer;
            oCompany.CompanyDB = BBDD;

#if DEBUG
            oCompany.DbServerType = SAPbobsCOM.BoDataServerTypes.dst_MSSQL2019;
#else
            oCompany.DbServerType = SAPbobsCOM.BoDataServerTypes.dst_HANADB;
#endif


            oCompany.DbUserName = userBBDD;
            oCompany.DbPassword = PassBBDD;
            oCompany.UserName = userSAP;
            oCompany.Password = PassSAP;
            oCompany.language = SAPbobsCOM.BoSuppLangs.ln_Spanish;
            oCompany.UseTrusted = false;

            int connectionResult = oCompany.Connect();

            if (connectionResult != 0)
            {
                int errorCode;
                String errorMessage;
                oCompany.GetLastError(out errorCode, out errorMessage);
                throw new Exception(errorCode.ToString() + errorMessage);
            }

            return oCompany;

        }

        public static bool ActualizaMargenes(SAPbobsCOM.Company oCompany, int docEntry, DataTable rows, ref string err)
        {
            SAPbobsCOM.Documents oDoc = null;
            string errMsg = string.Empty;
            int lretcode = 0;
            DataRow[] row;
            try
            {
                oDoc = (SAPbobsCOM.Documents)oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oOrders);

                if (oDoc.GetByKey(docEntry))
                {

                    bool need2update = false;
                    for (int i = 0; i < oDoc.Lines.Count; i++)
                    {
                        oDoc.Lines.SetCurrentLine(i);
                        row = rows.Select("LineNum = " + oDoc.Lines.LineNum.ToString());

                        if (rows.Rows.Count > 0)
                        {
                            //oDoc.Lines.GrossBase = -10; //Manual
                            oDoc.Lines.GrossBuyPrice = double.Parse(row[0]["PrecioCompra"].ToString());
                            //oDoc.Lines.GrossProfitTotalBasePrice = double.Parse(row[0]["PrecioCompra"].ToString()) * oDoc.Lines.Quantity;


                            need2update = true;
                        }

                    }
                    if (need2update)
                    {
                        lretcode = oDoc.Update();
                        if (lretcode != 0)
                        {
                            err = string.Format("Error actualizando documento, Nº {0} : {1}", oDoc.DocNum, oCompany.GetLastErrorDescription());
                            return false;
                        }

                    }
                }

                return true;
            }
            catch (Exception ex)
            {
                err = string.Format("Error actualizando documento, Nº {0} : {1}", docEntry, ex.Message);
                return false;
            }
            finally
            {
                System.Runtime.InteropServices.Marshal.ReleaseComObject(oDoc);
            }
        }

        public static bool CreaPersonaContacto(SAPbobsCOM.Company oCompany, string cardCode, string name, string Emails, bool isNew, ref string err)
        {
            SAPbobsCOM.BusinessPartners oBP = null;
            string errMsg = string.Empty;
            int lretcode = 0;
            string[] emails = Emails.Split(';');
            string EmailGroupCode = string.Empty;
            try
            {
                oBP = (SAPbobsCOM.BusinessPartners)oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oBusinessPartners);

                if (oBP.GetByKey(cardCode))
                {

                    bool need2update = false;

                    if (isNew)
                    {
                        if (oBP.ContactEmployees.Count > 1)
                        {
                            oBP.ContactEmployees.Add();
                        }
                        else
                        {
                            if (oBP.ContactEmployees.Name != "")
                                oBP.ContactEmployees.Add();
                        }

                        if (cUtils.EsEmailValido(Emails))
                        {
                            oBP.ContactEmployees.Name = name;
                            oBP.ContactEmployees.E_Mail = Emails;

                            need2update = true;
                        }
                        else
                        {
                            throw new Exception("email " + Emails + "no válido");
                        }
                        
                    }
                    else {
                        for (int i = 0; i < oBP.ContactEmployees.Count; i++)
                        {
                            oBP.ContactEmployees.SetCurrentLine(i);

                            if (oBP.ContactEmployees.Name == name)
                            {
                                oBP.ContactEmployees.E_Mail = emails[0];
                                EmailGroupCode = oBP.ContactEmployees.EmailGroupCode;
                                need2update = true;

                                for (int ii = 1; ii < emails.Length; ii++)
                                {
                                    if (!cUtils.EsEmailValido(emails[ii]))
                                    {
                                        oBP.ContactEmployees.Add();
                                        oBP.ContactEmployees.Name = name + ii.ToString();
                                        oBP.ContactEmployees.E_Mail = emails[ii];
                                        oBP.ContactEmployees.EmailGroupCode = EmailGroupCode;
                                    }
                                    else {
                                        throw new Exception("email " + emails[ii] + "no válido");
                                    }
                                    
                                }

                            }

                        }
                    }


                    
                    if (need2update)
                    {
                        lretcode = oBP.Update();
                        if (lretcode != 0)
                        {
                            err = string.Format("Error creando personas de contacto, cliente {0} : {1}", cardCode, oCompany.GetLastErrorDescription());
                            return false;
                        }

                    }
                }

                return true;
            }
            catch (Exception ex)
            {
                err = string.Format("Error creando personas de contacto, cliente {0} : {1}", cardCode, ex.Message);
                return false;
            }
            finally
            {
                System.Runtime.InteropServices.Marshal.ReleaseComObject(oBP);
            }
        }

        public static bool SendEmailViaMessageService(Logger Logger, Company oCompany, string cardCode, string to, string cc, string subject, string body, out string sErr, string[] attachmentPaths = null)
        {
            Message oMessage = null;
            MessagesService oMessageService = null;
            RecipientCollection oRecipientCollection = null;
            Attachments2 oAtt = null;
            string directory = string.Empty; //= System.IO.Path.GetDirectoryName(attachmentPath);        
            string fileName = string.Empty;//= System.IO.Path.GetFileNameWithoutExtension(attachmentPath); 
            string extension = string.Empty; //= System.IO.Path.GetExtension(attachmentPath)?.TrimStart('.');

            int lretcode = 0;
            int iAttach = 0;
            sErr = string.Empty;

            try
            {
                oMessageService = oCompany.GetCompanyService().GetBusinessService(SAPbobsCOM.ServiceTypes.MessagesService);
                oMessage = (SAPbobsCOM.Message)oMessageService.GetDataInterface(MessagesServiceDataInterfaces.msdiMessage);

                oMessage.Subject = subject;
                oMessage.Text = body;

                oRecipientCollection = oMessage.RecipientCollection;
                oRecipientCollection.Add();

                oRecipientCollection.Item(0).SendInternal = BoYesNoEnum.tNO;
                oRecipientCollection.Item(0).SendEmail = BoYesNoEnum.tYES;
                oRecipientCollection.Item(0).UserCode = cardCode;
                oRecipientCollection.Item(0).UserType = BoMsgRcpTypes.rt_ContactPerson;
                oRecipientCollection.Item(0).EmailAddress = to.ToString().ToLower();
               

                if (!string.IsNullOrEmpty(cc))
                {
                    oRecipientCollection.Add();

                    oRecipientCollection.Item(1).SendInternal = BoYesNoEnum.tNO;
                    oRecipientCollection.Item(1).SendEmail = BoYesNoEnum.tYES;
                    oRecipientCollection.Item(1).UserCode = cardCode;
                    oRecipientCollection.Item(1).UserType = BoMsgRcpTypes.rt_ContactPerson;
                    oRecipientCollection.Item(1).EmailAddress = to;
                }


                //if (!string.IsNullOrEmpty(attachmentPath) && System.IO.File.Exists(attachmentPath))
                //{
                //    oAtt = (SAPbobsCOM.Attachments2)oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oAttachments2);

                //    oAtt.Lines.SourcePath = directory;
                //    oAtt.Lines.FileName = fileName;
                //    oAtt.Lines.FileExtension = extension;

                //    lretcode = oAtt.Add();

                //    if (lretcode == 0)
                //    {
                //        iAttach = int.Parse(oCompany.GetNewObjectKey());
                //    }
                //    else
                //    {
                //        string err = string.Empty;
                //        int codeErr = 0;
                //        oCompany.GetLastError(out codeErr, out err);

                //        sErr = "Error creando doc. adjunto: " + err;
                //        return false;
                //    }

                //    oMessage.Attachment = iAttach;

                //}

                if (attachmentPaths != null && attachmentPaths.Length > 0)
                {
                    oAtt = (SAPbobsCOM.Attachments2)oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oAttachments2);

                    foreach (var attachmentPath in attachmentPaths)
                    {
                        if (!string.IsNullOrEmpty(attachmentPath) && System.IO.File.Exists(attachmentPath))
                        {
                            directory = System.IO.Path.GetDirectoryName(attachmentPath);
                            fileName = System.IO.Path.GetFileNameWithoutExtension(attachmentPath);
                            extension = System.IO.Path.GetExtension(attachmentPath)?.TrimStart('.');

                            oAtt.Lines.Add(); // Añadir nueva línea de adjunto
                            oAtt.Lines.SourcePath = directory;
                            oAtt.Lines.FileName = fileName;
                            oAtt.Lines.FileExtension = extension;
                        }
                    }

                    lretcode = oAtt.Add();

                    if (lretcode == 0)
                    {
                        iAttach = int.Parse(oCompany.GetNewObjectKey());
                        oMessage.Attachment = iAttach;
                    }
                    else
                    {
                        string err = string.Empty;
                        int codeErr = 0;
                        oCompany.GetLastError(out codeErr, out err);

                        sErr = "Error creando doc. adjunto: " + err;
                        return false;
                    }
                }

                oMessageService.SendMessage(oMessage);

                return true;
            }
            catch (Exception ex)
            {
                sErr = "Error enviando mensaje: " + ex.Message;
                return false;
            }
            finally
            {
                // Liberar objetos COM

                if (oAtt != null) System.Runtime.InteropServices.Marshal.ReleaseComObject(oAtt);
                if (oRecipientCollection != null) System.Runtime.InteropServices.Marshal.ReleaseComObject(oRecipientCollection);
                if (oMessage != null) System.Runtime.InteropServices.Marshal.ReleaseComObject(oMessage);
                if (oMessageService != null) System.Runtime.InteropServices.Marshal.ReleaseComObject(oMessageService);
            }
        }


        //public static void MandarMensaje(Company oCompany, string sKey, string Cliente)
        //{
        //    //DATOS DEL ENVIO DE MENSAJE
        //    SAPbobsCOM.Message oMessage = null;
        //    SAPbobsCOM.MessageDataColumns pMessageDataColumns = null;
        //    SAPbobsCOM.MessageDataColumn pMessageDataColumn = null;
        //    SAPbobsCOM.MessageDataLines oLines = null;
        //    SAPbobsCOM.MessageDataLine oLine = null;
        //    SAPbobsCOM.RecipientCollection oRecipientCollection = null;
        //    SAPbobsCOM.CompanyService oCmpSrv;
        //    SAPbobsCOM.MessagesService oMessageService;

        //    try
        //    {
        //        oCmpSrv = oCompany.GetCompanyService();
        //        //get msg service
        //        oMessageService = (SAPbobsCOM.MessagesService)oCmpSrv.GetBusinessService(SAPbobsCOM.ServiceTypes.MessagesService);
        //        oMessage = ((SAPbobsCOM.Message)(oMessageService.GetDataInterface(SAPbobsCOM.MessagesServiceDataInterfaces.msdiMessage)));

        //        //Asunto y texto de la notificación interna de SAP
        //        oMessage.Subject = "Se agregó una nueva cotización | " + sKey;
        //        oMessage.Text = "Se agregó la cotización número: " + sKey + ", del Cliente: " + Cliente + "\n\nPara ver los detalles oprime la flecha amarrilla debajo..";

        //        oRecipientCollection = oMessage.RecipientCollection;
        //        oRecipientCollection.Add();
        //        oRecipientCollection.Item(0).SendInternal = SAPbobsCOM.BoYesNoEnum.tYES;
        //        oRecipientCollection.Item(0).UserCode = "manager";
        //        pMessageDataColumns = oMessage.MessageDataColumns;
        //        pMessageDataColumn = pMessageDataColumns.Add();

        //        //Nombre de columna en el panel inferior de la notificación interna de SAP
        //        pMessageDataColumn.ColumnName = "Detalles";

        //        pMessageDataColumn.Link = SAPbobsCOM.BoYesNoEnum.tYES;
        //        //get lines
        //        oLines = pMessageDataColumn.MessageDataLines;
        //        //add new line
        //        oLine = oLines.Add();

        //        //Texto de la línea con el enlace al documento
        //        oLine.Value = "Enlace a Cotización: " + sKey;
        //        //Número de objeto cotización = 23
        //        oLine.Object = "13";
        //        //set the bo code
        //        oLine.ObjectKey = sKey;
        //        //Enviar mensaje
        //        oMessageService.SendMessage(oMessage);
        //    }
        //    catch (Exception ex)
        //    {

        //    }
        //}

        //public static bool GenPDFInvoice(Company oCompany, string layoutCode, string cardcode, int docentry, string filename)
        //{
        //    //SAPbobsCOM.ReportLayoutsService oLayoutService = (SAPbobsCOM.ReportLayoutsService)oCompany.GetCompanyService().GetBusinessService(SAPbobsCOM.ServiceTypes.ReportLayoutsService);
        //    //SAPbobsCOM.ReportParams oReportParams = (SAPbobsCOM.ReportParams)oLayoutService.GetDataInterface(SAPbobsCOM.ReportLayoutsServiceDataInterfaces.rlsdiReportParams);
        //    //oReportParams.ReportCode = "INV2";//defined in db table "RTYP"
        //    //oReportParams.CardCode = "C20000";//business partner 
        //    //var oReport = oLayoutService.GetDefaultReport(oReportParams);
        //    //BlobParams oBlobParams = (SAPbobsCOM.BlobParams)oCompany.GetCompanyService().GetDataInterface(SAPbobsCOM.CompanyServiceDataInterfaces.csdiBlobParams);
        //    //oBlobParams.Table = "RDOC";
        //    //oBlobParams.Field = "Template";
        //    //oBlobParams.FileName = @"C:\salesorder.rpt";
        //    //BlobTableKeySegment oKeySegment = oBlobParams.BlobTableKeySegments.Add();
        //    //oKeySegment.Name = "DocCode";
        //    //oKeySegment.Value = oReport.LayoutCode;
        //    //SBO_Company.GetCompanyService().SaveBlobToFile(oBlobParams);


        //    //SAPbobsCOM.CompanyService oCmpSrv;
        //    //SAPbobsCOM.ReportLayoutsService oReportLayoutService;
        //    //SAPbobsCOM.ReportLayoutPrintParams oPrintParam;

        //    //oCmpSrv = oCompany.GetCompanyService();

        //    //oReportLayoutService = (SAPbobsCOM.ReportLayoutsService)oCmpSrv.GetBusinessService(SAPbobsCOM.ServiceTypes.ReportLayoutsService);

        //    //oPrintParam = (SAPbobsCOM.ReportLayoutPrintParams)
        //    //    oReportLayoutService.GetDataInterface(
        //    //        SAPbobsCOM.ReportLayoutsServiceDataInterfaces.rlsdiReportLayoutPrintParams
        //    //    );

        //    //oPrintParam.LayoutCode = layoutCode;
        //    //oPrintParam.DocEntry = docentry;

        //    //oReportLayoutService.Print(oPrintParam);

        //    ReportDocument cryReportDocument = new ReportDocument();
        //    cryReportDocument.Load(sPlantilla);
        //    cryReportDocument.SetDatabaseLogon(usuarioBD, passBD, servidorBD, nombreBD);
        //    cryReportDocument.SetParameterValue("DocKey@", iDocEntry);
        //    cryReportDocument.SetParameterValue("Cadena@", Cadena);
        //    cryReportDocument.SetParameterValue("Firma@", Firma);
        //    cryReportDocument.ExportToDisk(ExportFormatType.PortableDocFormat, this.AttachPDF);
        //    cryReportDocument.Dispose();
        //    EventLog.WriteEntry(sSource, "XML y PDF Creados");

        //}



        public static void ExportPDFCrystalReport(Logger Logger, String rptPath, Dictionary<String, object> paramters, String pathPdf = "")
        {
            ReportDocument rpt = null;

            try
            {
                if (pathPdf == "")
                    pathPdf = System.IO.Path.GetTempPath() + System.IO.Path.GetFileName(rptPath) + ".pdf";

                rpt = new ReportDocument();
                rpt.Load(rptPath);

                string strConnection = string.Format("DRIVER={0};UID={1};PWD={2};SERVERNODE={3};DATABASE={4};", Properties.Settings.Default.DriverCR,
                    Properties.Settings.Default.UserBBDD, EncryptionHelper.DecryptPass(Properties.Settings.Default.PwdBBDD), Properties.Settings.Default.ServerIP, Properties.Settings.Default.BBDD);


                CrystalDecisions.Shared.NameValuePairs2 logonProps2 = rpt.DataSourceConnections[0].LogonProperties;

                logonProps2.Set("Provider", Properties.Settings.Default.DriverCR);
                logonProps2.Set("Server Type", Properties.Settings.Default.DriverCR);
                logonProps2.Set("Connection String", strConnection);

                rpt.DataSourceConnections.Clear();

                rpt.DataSourceConnections[0].SetLogonProperties(logonProps2);
                rpt.DataSourceConnections[0].SetConnection(Properties.Settings.Default.ServerIP, Properties.Settings.Default.BBDD, false);

                rpt.Refresh();

                foreach (var item in paramters)
                {
                    rpt.SetParameterValue(item.Key, item.Value);
                }

                rpt.ExportToDisk(CrystalDecisions.Shared.ExportFormatType.PortableDocFormat, pathPdf);

            }
            catch (Exception ex)
            {

                throw new Exception("Error generando pdf: " + ex.Message);
            }
            finally
            {
                if (rpt != null)
                {
                    rpt.Dispose();
                    rpt = null;
                }

            }

        }

        public void Dispose()
        {
            GC.SuppressFinalize(this);
            GC.Collect();
        }
    }
}

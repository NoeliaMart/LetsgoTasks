using System;
using System.Collections.Generic;
using System.Data;
using LetsGoTasks.Utils.Models;

namespace LetsGoTasks.Utils
{
    public class cDIAPI
    {
        public static SAPbobsCOM.Company ConnectionCompany(string companyDB)
        {
            var oCompany = new SAPbobsCOM.Company();

            oCompany.Server = Properties.Settings.Default.ServerSAP;
            oCompany.SLDServer = Properties.Settings.Default.SLDServer;
            oCompany.CompanyDB = companyDB;

#if DEBUG
            oCompany.DbServerType = SAPbobsCOM.BoDataServerTypes.dst_MSSQL2019;
#else
            oCompany.DbServerType = SAPbobsCOM.BoDataServerTypes.dst_HANADB;
#endif

            oCompany.DbUserName = Properties.Settings.Default.UserBBDD;
            oCompany.DbPassword = EncryptionHelper.DecryptPass(Properties.Settings.Default.PwdBBDD);
            oCompany.UserName = Properties.Settings.Default.UserSAP;
            oCompany.Password = EncryptionHelper.DecryptPass(Properties.Settings.Default.PwdSAP);
            oCompany.language = SAPbobsCOM.BoSuppLangs.ln_Spanish;
            oCompany.UseTrusted = false;

            if (oCompany.Connect() != 0)
            {
                oCompany.GetLastError(out int code, out string message);
                throw new Exception($"{code} {message}");
            }

            return oCompany;
        }

        public static bool CreaProveedor(SAPbobsCOM.Company company, RegistroFactura r, ref string err, ref string cardCode)
        {
            SAPbobsCOM.BusinessPartners oBP = null;

            try
            {
                oBP = (SAPbobsCOM.BusinessPartners)company.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oBusinessPartners);

                oBP.CardType = r.CardType == "cSupplier" ? SAPbobsCOM.BoCardTypes.cSupplier : SAPbobsCOM.BoCardTypes.cLid;
                oBP.FederalTaxID = r.LicTradNum;
                oBP.Series = r.Series;
                oBP.CardName = r.CardName;
                oBP.Address = r.Address;
                oBP.City = r.City;
                oBP.ZipCode = r.ZipCode;
                oBP.Country = r.Country;
                oBP.LanguageCode = 23;

                if (!string.IsNullOrEmpty(r.WTCode))
                {
                    oBP.SubjectToWithholdingTax = SAPbobsCOM.BoYesNoNoneEnum.boYES;
                    oBP.WTCode = r.WTCode;
                }

                if (!string.IsNullOrEmpty(r.PeymentMethodCode))
                {
                    oBP.BPPaymentMethods.PaymentMethodCode = r.PeymentMethodCode;
                    oBP.BPPaymentMethods.Add();

                    oBP.PeymentMethodCode = r.PeymentMethodCode;
                }
                
                //bp.PayTermsGrpCode = r.PayTermsGrpCode;

                if (oBP.Add() != 0)
                {
                    err = $"Error procesando proveedor {r.LicTradNum}: {company.GetLastErrorDescription()}";
                    return false;
                }

                cardCode = company.GetNewObjectKey();
                return true;
            }
            catch (Exception ex)
            {
                err = $"Error procesando proveedor {r.LicTradNum}: {ex.Message}";
                return false;
            }
            finally
            {
                if (oBP != null)
                    System.Runtime.InteropServices.Marshal.ReleaseComObject(oBP);
            }
        }

        public static bool CreaFacturaCompra(SAPbobsCOM.Company company,string cardCode,List<RegistroFactura> lineas,ref string err, ref int docEntry)
        {
            SAPbobsCOM.Documents oDoc = null;
            bool primera_linea = true;
            var r = lineas[0];
            try
            {
                if (lineas == null || lineas.Count == 0)
                {
                    err = "La factura no contiene líneas.";
                    return false;
                }

                oDoc = (SAPbobsCOM.Documents)company.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oPurchaseInvoices);

                oDoc.CardCode = cardCode;
                oDoc.DocType = SAPbobsCOM.BoDocumentTypes.dDocument_Service;
                oDoc.NumAtCard = r.NumAtCard;
                oDoc.DocCurrency = r.DocCurrency;

                if (!string.IsNullOrEmpty(r.PaymentMethod)) oDoc.PaymentMethod = r.PaymentMethod;

                //oDoc.PaymentGroupCode = r.PaymentGroupCode;
                oDoc.DocDate = r.DocDate;
                oDoc.TaxDate = r.TaxDate;
                oDoc.DocDueDate = r.DocDueDate;

                if (!string.IsNullOrEmpty(r.WTCode)) oDoc.WithholdingTaxData.WTCode = r.WTCode;


                foreach (var l in lineas)
                {
                    if (!primera_linea)
                    {
                        oDoc.Lines.Add();
                    }

                    oDoc.Lines.ItemDescription = l.Description;
                    oDoc.Lines.AccountCode = l.AccountCode;
                    oDoc.Lines.LineTotal = l.LineTotal;
                    oDoc.Lines.VatGroup = l.VatGroup;
                    oDoc.Lines.UserFields.Fields.Item("U_Observaciones").Value = l.U_Observaciones;
                    oDoc.Lines.UserFields.Fields.Item("U_COD_Ocs").Value = l.U_COD_Ocs;
                    oDoc.Lines.UserFields.Fields.Item("U_Linea_COD_Ocs").Value = l.U_Linea_COD_Ocs;
                    oDoc.Lines.UserFields.Fields.Item("U_Archivo_ad").Value = l.Archivo_adjunto;

                    if (l.CostingCode != string.Empty)
                        oDoc.Lines.CostingCode = l.CostingCode;

                    if (l.CostingCode != string.Empty)
                        oDoc.Lines.CostingCode2 = l.CostingCode2;

                    if (l.CostingCode != string.Empty)
                        oDoc.Lines.CostingCode3 = l.CostingCode3;

                    if (l.ProjectCode != string.Empty)
                        oDoc.Lines.ProjectCode = l.ProjectCode;


                    primera_linea = false;
                }

                if (oDoc.Add() != 0)
                {
                    err = $"Error factura {r.NumAtCard}: {company.GetLastErrorDescription()}";
                    return false;
                }

                docEntry = int.Parse(company.GetNewObjectKey());
                return true;
            }
            catch (Exception ex)
            {
                err = $"Error factura {r.NumAtCard}: {ex.Message}";
                return false;
            }
            finally
            {
                if (oDoc != null)
                    System.Runtime.InteropServices.Marshal.ReleaseComObject(oDoc);
            }
        }
    }
}

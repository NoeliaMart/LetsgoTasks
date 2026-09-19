using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace LetsGoTasks.Utils.Models
{
    // ---------------------------------------------------------
    // MODELO DE DATOS
    // ---------------------------------------------------------
    public class RegistroFactura
    {
        // PROVEEDOR
        public string CardCode { get; set; }
        public string CardName { get; set; }
        public string LicTradNum { get; set; }
        public int Series { get; set; }
        public string Address { get; set; }
        public string City { get; set; }
        public string ZipCode { get; set; }
        public string Country { get; set; }
        public string PeymentMethodCode { get; set; }
        public int PayTermsGrpCode { get; set; }
        public string CardType { get; set; }
        public string WTCode { get; set; }
        public string State { get; set; }

        // FACTURA
        public string NumAtCard { get; set; }
        public DateTime DocDate { get; set; }
        public DateTime DocDueDate { get; set; }
        public DateTime TaxDate { get; set; }
        public string DocCurrency { get; set; }
        public string PaymentMethod { get; set; }
        public int PaymentGroupCode { get; set; }
        public string U_COD_Ocs { get; set; }
        

        // LÍNEA
        public string AccountCode { get; set; }
        public string VatGroup { get; set; }
        public string Description { get; set; }
        public double LineTotal { get; set; }
        public string U_Observaciones { get; set; }
        public string U_Linea_COD_Ocs { get; set; }
        public string CostingCode { get; set; }
        public string CostingCode2 { get; set; }
        public string CostingCode3 { get; set; }
        public string ProjectCode { get; set; }
        public string Archivo_adjunto { get; set; }
    }
}

using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace PDFToExcel
{
    public class LiniaLiquidacio
    {
        public string? CIA { get; set; }
        public string? TRNC { get; set; }
        public string? NUM_DOC { get; set; }
        public string? FECHA_EMISION { get; set; }
        public string? CPUI { get; set; }
        public string? NR_CODE { get; set; }
        public string? STAT { get; set; }
        public string? FOP { get; set; }
        public double? IMPORT_TRANSACC { get; set; }
        public double? TARIFA { get; set; }
        public double? TASAS { get; set; }
        public double? G_C { get; set; }
        public double? PEN { get; set; }
        public double? COBL { get; set; }
        public double? STD_PERC { get; set; }
        public double? STD_IMPORTE { get; set; }
        public double? SUPP_PERC { get; set; }
        public double? SUPP_IMPORTE { get; set; }

        public double? IVA_S_COM { get; set; }
        public double? NETO_PAG { get; set; }
    }
}

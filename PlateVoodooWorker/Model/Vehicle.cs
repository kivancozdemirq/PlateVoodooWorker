using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace PlateVoodooWorker.Model
{
    public class Vehicle
    {
        public int Id { get; set; }
        public string TransactionId { get; set; }
        public string? FrontPlate { get; set; }
        public string? FrontCoordination { get; set; }
        public double? FrontConf { get; set; }
        public string? RearPlate { get; set; }
        public string? RearCoordination { get; set; }
        public double? RearConf { get; set; }
        public string? UpdateDate { get; set; }
        public string? Description { get; set; }
        public int? State { get; set; }
        public string? Aselsan_REAR { get; set; }
        public decimal? Aselsan_REARCONF { get; set; }
        public string? Aselsan_FRONT { get; set; }
        public decimal? Aselsan_FRONTCONF { get; set; }
        public string? Aselsan_FINAL { get; set; }
        public decimal? Aselsan_FINALCONF { get; set; }
        public string? Aselsan_APPROVEDPLATENUMBER { get; set; }
        public string? Aselsan_APPROVEDCLASS { get; set; }
        public string? Aselsan_TABLETYPE { get; set; }
        public int? OperationId { get; set; }
    }
}

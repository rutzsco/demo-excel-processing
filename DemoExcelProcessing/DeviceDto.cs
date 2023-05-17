using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace DemoExcelProcessing
{
    public class DeviceDto
    {
        public int Id { get; set; }
        public string DeviceId { get; set; }
        public string DeviceName { get; set; }
        public string DeviceType { get; set; }
        public int? MappedSpaceId { get; set; }
        public string MappedSpaceName { get; set; }
        public string IpAddress { get; set; }
    }
}

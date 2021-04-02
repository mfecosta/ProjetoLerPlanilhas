using System;
using System.Collections.Generic;
using System.ComponentModel.DataAnnotations;
using System.Linq;
using System.Threading.Tasks;

namespace ExecutarScripts.Models
{
    public class StaffInfoViewModel
    {
        [Required(ErrorMessage ="Carrega a planilha !")]
        public string Scripts { get; set; }
        public List<StaffInfoViewModel> StaffList { get; set; }

        public StaffInfoViewModel()
        {
            StaffList = new List<StaffInfoViewModel>();
        }
    }
}

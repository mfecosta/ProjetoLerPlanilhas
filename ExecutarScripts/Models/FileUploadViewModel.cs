using Microsoft.AspNetCore.Http;
using System;
using System.Collections.Generic;
using System.ComponentModel.DataAnnotations;
using System.Linq;
using System.Threading.Tasks;

namespace ExecutarScripts.Models
{
    public class FileUploadViewModel
    {
        [Required(ErrorMessage = "Carrega a planila !")]
        public IFormFile Planilha { get; set; }
        [Required(ErrorMessage = "Carrega a planila !")]
        public StaffInfoViewModel StaffInfoViewModel { get; set; }

        public FileUploadViewModel()
        {
            StaffInfoViewModel = new StaffInfoViewModel();
        }
    }
}

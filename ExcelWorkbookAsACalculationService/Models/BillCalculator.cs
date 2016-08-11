//Copyright (c) CodeMoggy. All rights reserved. Licensed under the MIT license.
//See LICENSE in the project root for license information.

using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.ComponentModel.DataAnnotations;
using System.Linq;
using System.Web;

namespace ExcelWorkbookAsACalculationService.Models
{
    public class BillCalculator
    {
        [Required]
        [Display(Name="Bill Amount")]
        [DisplayFormat(DataFormatString = "{0:n2}", ApplyFormatInEditMode = true)]
        public decimal BillAmount { get; set; }
        [Required]
        [Display(Name="Tip (%)")]
        [DisplayFormat(DataFormatString = "{0:n2}", ApplyFormatInEditMode = true)]
        public decimal TipPercentage { get; set; }
        [Required]
        [Display(Name="Number in Party")]
        public int NumberOfPeople { get; set; }
        [Required]
        [Display(Name="Total Bill (incl. tip)")]
        [DisplayFormat(DataFormatString = "{0:n2}", ApplyFormatInEditMode = true)]
        public decimal BillPlusTip { get; set; }
        [Required]
        [Display(Name="Amount per Person")]
        [DisplayFormat(DataFormatString = "{0:n2}", ApplyFormatInEditMode = true)]
        public decimal AmountPerPerson { get; set; }

        public BillCalculator() { }

        public BillCalculator(
            decimal billAmount,
            decimal tipPercentage,
            int numberOfPeople,
            decimal billPlusTip,
            decimal amountPerPerson)
        {
            BillAmount = billAmount;
            TipPercentage = tipPercentage;
            NumberOfPeople = numberOfPeople;
            BillPlusTip = billPlusTip;
            AmountPerPerson = amountPerPerson;
        }
    }
}
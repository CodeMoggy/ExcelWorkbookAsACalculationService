﻿//Copyright (c) CodeMoggy. All rights reserved. Licensed under the MIT license.
//See LICENSE in the project root for license information.

using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;

namespace ExcelWorkbookAsACalculationService.Models
{
    public class ValuesRequest : IExcelRequest
    {
        public List<List<string>> values { get; set; }
    }
}
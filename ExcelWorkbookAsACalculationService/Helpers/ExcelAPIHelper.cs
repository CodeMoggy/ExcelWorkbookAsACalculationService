//Copyright (c) CodeMoggy. All rights reserved. Licensed under the MIT license.
//See LICENSE in the project root for license information.

using ExcelWorkbookAsACalculationService.Models;
using Newtonsoft.Json;
using Newtonsoft.Json.Linq;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Net.Http;
using System.Net.Http.Headers;
using System.Text;
using System.Threading.Tasks;
using System.Web;

namespace ExcelWorkbookAsACalculationService.Helpers
{
    public class ExcelAPIHelper
    {
        /// <summary>
        /// gets the nameditem values from the Excel workbook
        /// to reduce the number of requests try considering the nameditem as being a collection of cells rather than a nameditem per cell that I have here
        /// </summary>
        /// <param name="billCalculatorModel"></param>
        /// <param name="baseWorkbookUrl"></param>
        /// <param name="accessToken"></param>
        /// <param name="sessionId"></param>
        /// <returns></returns>
        public async Task<BillCalculator> LoadValuesFromWorkbookAsync(BillCalculator billCalculatorModel, string baseWorkbookUrl, string accessToken, string sessionId)
        {
            // get the loanMoadel values from the Excel document.
            billCalculatorModel.TipPercentage = Decimal.Round(Decimal.Parse(await ParseNamedValueExcelRequest("TipAsPercent", baseWorkbookUrl, accessToken, sessionId)) * 100, 2);
            billCalculatorModel.BillAmount = Decimal.Round(Decimal.Parse(await ParseNamedValueExcelRequest("BillAmount", baseWorkbookUrl, accessToken, sessionId)), 2);
            billCalculatorModel.NumberOfPeople = int.Parse(await ParseNamedValueExcelRequest("NumberOfPeople", baseWorkbookUrl, accessToken, sessionId));
            billCalculatorModel.BillPlusTip = Decimal.Round(Decimal.Parse(await ParseNamedValueExcelRequest("BillPlusTip", baseWorkbookUrl, accessToken, sessionId)), 2);
            billCalculatorModel.AmountPerPerson = Decimal.Round(Decimal.Parse(await ParseNamedValueExcelRequest("AmountPerPerson", baseWorkbookUrl, accessToken, sessionId)), 2);

            return billCalculatorModel;
        }

        /// <summary>
        /// Based on the nameditem and value - update the Excel workbook
        /// Like the get, try considering having a single nameditem and set the values in a single request
        /// </summary>
        /// <param name="name"></param>
        /// <param name="namedValue"></param>
        /// <param name="baseWorkbookUrl"></param>
        /// <param name="accessToken"></param>
        /// <param name="sessionId"></param>
        /// <returns></returns>
        public async Task SetNamedValueExcelRequestAsync(string name, string namedValue, string baseWorkbookUrl, string accessToken, string sessionId)
        {
            // set a namedItem in the document and return the calculated value for the document
            var restUrl = baseWorkbookUrl + string.Format("names('{0}')/Range", name);
            var result = await DoExcelRequestAsync(HttpMethod.Get, null, restUrl, accessToken, sessionId);

            // get the address of the nameditem 
            dynamic i = JObject.Parse(result);
            string address = i.address;

            // address is of the format worksheet!cell
            var worksheet = address.Split('!')[0];
            var cellAddress = address.Split('!')[1];

            // update the worksheet cell with the new value using Patch
            restUrl = baseWorkbookUrl + string.Format("worksheets({0})/range(address='{1}')", worksheet, cellAddress);

            var value = new List<string>() { namedValue };
            var values = new List<List<string>>() { value };
            var valuesRequest = new ValuesRequest { values = values };

            result = await DoExcelRequestAsync(new HttpMethod("PATCH"), valuesRequest, restUrl, accessToken, sessionId);

        }

        /// <summary>
        /// Creates a new session to enable new values to be persisted to the workbook
        /// </summary>
        /// <param name="baseWorkbookUrl"></param>
        /// <param name="accessToken"></param>
        /// <returns></returns>
        public async Task<string> GetWorkbookSessionIdAsync(string baseWorkbookUrl, string accessToken)
        {
            // create a new workbook session
            var restUrl = baseWorkbookUrl + "createSession";
            var sessionRequest = new SessionRequest { persistChanges = "true" };
            var result = await DoExcelRequestAsync(HttpMethod.Post, sessionRequest, restUrl, accessToken, string.Empty);
            // get the sessionId from the result
            dynamic d = JObject.Parse(result);
            string sessionId = d.id;

            return sessionId;
        }

        /// <summary>
        /// Based on the nameditem get is value
        /// </summary>
        /// <param name="name"></param>
        /// <param name="baseWorkbookUrl"></param>
        /// <param name="accessToken"></param>
        /// <param name="sessionId"></param>
        /// <returns></returns>
        private async Task<string> ParseNamedValueExcelRequest(string name, string baseWorkbookUrl, string accessToken, string sessionId)
        {
            // InterestRate
            // set a namedItem in the document and return the calculated value for the document
            var restUrl = baseWorkbookUrl + string.Format(@"names('{0}')/range", name);
            var result = await DoExcelRequestAsync(HttpMethod.Get, null, restUrl, accessToken, sessionId);

            
            var jsonResult = JObject.Parse(result);

            string value = string.Empty;

            if (name.Equals("NumberOfPeople"))
                value = jsonResult["values"][0][0].ToString();
            else if (name.Equals("TipAsPercent"))
                value = string.Format("{0:f4}", jsonResult["values"][0][0]);
            else
                value = string.Format("{0:f2}", jsonResult["values"][0][0]);





            return value;
        }

        /// <summary>
        /// Based on the ExcelRequest and the HttpMethod make the request to the Excel REST API
        /// </summary>
        /// <param name="method"></param>
        /// <param name="excelRequest"></param>
        /// <param name="url"></param>
        /// <param name="accessToken"></param>
        /// <param name="sessionId"></param>
        /// <returns></returns>
        private async Task<string> DoExcelRequestAsync(HttpMethod method, IExcelRequest excelRequest, string url, string accessToken, string sessionId)
        {
            string result = string.Empty;

            using (var client = new HttpClient())
            {
                using (var request = new HttpRequestMessage(method, url))
                {
                    request.Headers.Authorization = new AuthenticationHeaderValue("Bearer", accessToken);

                    if (!string.IsNullOrEmpty(sessionId))
                        request.Headers.Add("workbook-session-id", sessionId);

                    if (excelRequest != null)
                    {
                        request.Content = new StringContent(JsonConvert.SerializeObject(excelRequest), Encoding.UTF8, "application/json");
                    }

                    using (HttpResponseMessage response = await client.SendAsync(request))
                    {
                        if (response.IsSuccessStatusCode)
                        {
                            result = await response.Content.ReadAsStringAsync();
                        }
                        else
                        {
                            var reason = response.ReasonPhrase;
                        }
                    }
                }
            }

            return result;
        }
    }
}
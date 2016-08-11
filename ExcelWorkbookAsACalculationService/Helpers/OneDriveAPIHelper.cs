//Copyright (c) CodeMoggy. All rights reserved. Licensed under the MIT license.
//See LICENSE in the project root for license information.

using Newtonsoft.Json.Linq;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Net.Http;
using System.Net.Http.Headers;
using System.Threading.Tasks;
using System.Web;

namespace ExcelWorkbookAsACalculationService.Helpers
{
    public class OneDriveAPIHelper
    {
        /// <summary>
        /// If the Excel workbook does not exist in the User's OneDrive then this method uploads the workbook from the Assets folder.
        /// It returns the base URL where the Excel workbook was uploaded to.
        /// </summary>
        /// <param name="accessToken"></param>
        /// <returns></returns>
        public async Task<string> CreateFileIfNotExistsAsync(string accessToken)
        {
            string fileName = "RestaurantBillCalculator.xlsx";
            string excelUrlBase = string.Empty;
            string baseSearchFileEndpoint = "https://graph.microsoft.com/v1.0/me/drive/root/children/";
            string baseFileEndpoint = "https://graph.microsoft.com/v1.0/me/drive/items/";
            string fileId = string.Empty;

            // read the OneDrive and find the item
            using (var client = new HttpClient())
            {
                using (var request = new HttpRequestMessage(HttpMethod.Get, baseSearchFileEndpoint + "?$select=name,id"))
                {
                    request.Headers.Authorization = new AuthenticationHeaderValue("Bearer", accessToken);

                    using (HttpResponseMessage response = await client.SendAsync(request))
                    {
                        if (response.IsSuccessStatusCode)
                        {
                            var result = await response.Content.ReadAsStringAsync();

                            JObject parsedResult = JObject.Parse(result);

                            // determine if the Excel workbook exists in the user's OneDrive
                            foreach (JObject file in parsedResult["value"])
                            {
                                var name = (string)file["name"];
                                if (name.Contains(fileName))
                                {
                                    // workbook file found
                                    fileId = (string)file["id"];
                                    excelUrlBase = baseFileEndpoint + fileId + "/workbook/";
                                }
                            }
                        }
                        else
                        {
                            // handle error response
                        }
                    }

                }

                if (string.IsNullOrEmpty(excelUrlBase))
                {
                    // didn't find the file so upload the workbook RestaurantBillCalculator.xlsx in the Assets floder to the root folder of the user's OneDrive

                    string absPath = System.Web.HttpContext.Current.Server.MapPath("Assets/" + fileName);

                    var excelFile = System.IO.File.OpenRead(absPath);
                    byte[] contents = new byte[excelFile.Length];
                    excelFile.Read(contents, 0, (int)excelFile.Length);
                    excelFile.Close();
                    var contentStream = new MemoryStream(contents);

                    var contentPostBody = new StreamContent(contentStream);
                    contentPostBody.Headers.Add("Content-Type", "application/octet-stream");

                    using (var request = new HttpRequestMessage(HttpMethod.Put, baseSearchFileEndpoint + "/" + fileName + "/content"))
                    {
                        request.Headers.Authorization = new AuthenticationHeaderValue("Bearer", accessToken);
                        request.Content = contentPostBody;

                        using (HttpResponseMessage response = await client.SendAsync(request))
                        {
                            if (response.IsSuccessStatusCode)
                            {
                                //Get the Id of the new file.
                                var responseContent = await response.Content.ReadAsStringAsync();
                                var parsedResponse = JObject.Parse(responseContent);
                                fileId = (string)parsedResponse["id"];
                                excelUrlBase = baseFileEndpoint + fileId + "/workbook/";
                            }
                            else
                            {
                                // handle error response
                            }
                        }

                    }
                }
            }

            return excelUrlBase;

        }
    }
}
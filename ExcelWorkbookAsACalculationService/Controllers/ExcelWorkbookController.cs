//Copyright (c) CodeMoggy. All rights reserved. Licensed under the MIT license.
//See LICENSE in the project root for license information.

using ExcelWorkbookAsACalculationService.Helpers;
using ExcelWorkbookAsACalculationService.Models;
using System;
using System.Threading.Tasks;
using System.Web;
using System.Web.Mvc;
using System.Configuration;
using System.Security.Claims;
using Microsoft.IdentityModel.Clients.ActiveDirectory;
using Microsoft.Owin.Security;
using Microsoft.Owin.Security.OpenIdConnect;

namespace ExcelWorkbookAsACalculationService.Controllers
{
    [Authorize]
    public class ExcelWorkbookController : Controller
    {
        private string clientId = ConfigurationManager.AppSettings["ida:ClientId"];
        private string appKey = ConfigurationManager.AppSettings["ida:ClientSecret"];
        private string aadInstance = ConfigurationManager.AppSettings["ida:AADInstance"];
        private ExcelAPIHelper excelHelper = new ExcelAPIHelper();
        private OneDriveAPIHelper oneDriveHelper = new OneDriveAPIHelper();


        /// <summary>
        /// When the page is launched, go and get the current values from the Excel workbook 
        /// </summary>
        /// <returns></returns>
        // GET: ExcelWorkbook
        public async Task<ActionResult> Index()
        {
            string accessToken = string.Empty;

            try
            {
                // get a valid aad token
                accessToken = await GetGraphAccessTokenAsync();

                // if the workbook does not exist on the user's OneDrive then upload the book (see Assets folder)
                var baseWorkbookUrl = await oneDriveHelper.CreateFileIfNotExistsAsync(accessToken);

                // get a session for the requests - this is important if the values are to be persisted
                var sessionId = await excelHelper.GetWorkbookSessionIdAsync(baseWorkbookUrl, accessToken);

                // create a new model (this is the data binding for the page)
                BillCalculator billCalculatorModel = new BillCalculator();

                // get the current values from the workbook
                billCalculatorModel = await excelHelper.LoadValuesFromWorkbookAsync(billCalculatorModel, baseWorkbookUrl, accessToken, sessionId);

                return View("Index", billCalculatorModel);
            }
            catch (AdalException)
            {
                // Return to error page.
                return View("Error");
            }
            // if the above failed, the user needs to explicitly re-authenticate for the app to obtain the required token
            catch (Exception)
            {
                return View("Relogin");
            }
        }

        /// <summary>
        /// Provides the necessary logic to determine the data that needs to be passed to the Excel REST API
        /// </summary>
        /// <param name="billCalculatorModel"></param>
        /// <returns></returns>
        [HttpPost]
        public async Task<ActionResult> Calculate(BillCalculator billCalculatorModel)
        {
            string accessToken = string.Empty;

            try
            {
                // grab the AAD access token
                accessToken = await GetGraphAccessTokenAsync();

                // does workbook exist in OneDrive
                var baseWorkbookUrl = await oneDriveHelper.CreateFileIfNotExistsAsync(accessToken);

                // get a session for the requests - this is important if the values are to be persisted
                var sessionId = await excelHelper.GetWorkbookSessionIdAsync(baseWorkbookUrl, accessToken);

                // these are the values the user can edit so pass these to Excel workbook
                await excelHelper.SetNamedValueExcelRequestAsync("BillAmount", billCalculatorModel.BillAmount.ToString(), baseWorkbookUrl, accessToken, sessionId);
                await excelHelper.SetNamedValueExcelRequestAsync("TipAsPercent", (billCalculatorModel.TipPercentage / 100).ToString(), baseWorkbookUrl, accessToken, sessionId);
                await excelHelper.SetNamedValueExcelRequestAsync("NumberOfPeople", billCalculatorModel.NumberOfPeople.ToString(), baseWorkbookUrl, accessToken, sessionId);

                // clear the current model state
                ModelState.Clear();

                // get the new values from the Excel workbook
                billCalculatorModel = await excelHelper.LoadValuesFromWorkbookAsync(billCalculatorModel, baseWorkbookUrl, accessToken, sessionId);

                return View("Index", billCalculatorModel);
            }
            catch (AdalException)
            {
                // Return to error page.
                return View("Error");
            }
            // if the above failed, the user needs to explicitly re-authenticate for the app to obtain the required token
            catch (Exception)
            {
                return View("Relogin");
            }

        }
        

        private async Task<string> GetGraphAccessTokenAsync()
        {
            var signInUserId = ClaimsPrincipal.Current.FindFirst(ClaimTypes.NameIdentifier).Value;
            var userObjectId = ClaimsPrincipal.Current.FindFirst("http://schemas.microsoft.com/identity/claims/objectidentifier").Value;
            string tenantID = ClaimsPrincipal.Current.FindFirst("http://schemas.microsoft.com/identity/claims/tenantid").Value;
            string authority = aadInstance + tenantID;
            SessionTokenCache tokenCache = new SessionTokenCache(userObjectId, HttpContext);

            AuthHelper authHelper = new AuthHelper(authority, clientId, appKey, tokenCache);
            string accessToken = await authHelper.GetUserAccessToken(Url.Action("Index", "Home", null, Request.Url.Scheme));

            return accessToken;
        }

        public void RefreshSession()
        {
            HttpContext.GetOwinContext().Authentication.Challenge(
                new AuthenticationProperties { RedirectUri = "/ExcelWorkbook" },
                OpenIdConnectAuthenticationDefaults.AuthenticationType);
        }


    }
}
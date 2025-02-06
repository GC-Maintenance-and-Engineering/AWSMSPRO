using System;
using System.Collections.Generic;
using System.Linq;
using System.Net.Http;
using System.Net.Http.Headers;
using System.Text;
using System.Threading.Tasks;
using System.Web;

namespace MeExpert.CommonProvide
{
    public class clsApi
    {
        /*//https://localhost:44306/  wApiCenter/api/  */
        private string strApiKey = "asdqw%3&*$$rfcWrve1433sdfAR470m9lWGFG";
        private string strRightKey = "ApiCenter:SwSe#Srwer453234ppisdfkSFSDs)&*&^nsSGSw[w0!";
        private string strUriApi = "http://gcgpgcmeapp01/";
        private string strDefFn = " wApiCenter/api/";
    

        public async Task<HttpResponseMessage> wsGet(string strFn)
        {
            HttpResponseMessage hrmResult = new HttpResponseMessage();
            HttpClient hpcClient = new HttpClient();

            hpcClient.BaseAddress = new Uri(strUriApi);
            hpcClient.DefaultRequestHeaders.Add("API_KEY", strApiKey);

            var byteArray = new UTF8Encoding().GetBytes(strRightKey);
            hpcClient.DefaultRequestHeaders.Authorization = new AuthenticationHeaderValue("Basic", Convert.ToBase64String(byteArray));
            hpcClient.DefaultRequestHeaders.Accept.Add(new MediaTypeWithQualityHeaderValue("application/json"));
            hrmResult = await hpcClient.GetAsync(strDefFn + strFn);

            return hrmResult;

        }

        public async Task<HttpResponseMessage> wsPost(string strFn , string strJsonParameter)
        {
            HttpResponseMessage hrmResult = new HttpResponseMessage();
            HttpClient hpcClient = new HttpClient();

            hpcClient.BaseAddress = new Uri(strUriApi);           
            hpcClient.DefaultRequestHeaders.Add("API_KEY", strApiKey);

            var byteArray = new UTF8Encoding().GetBytes(strRightKey);
            hpcClient.DefaultRequestHeaders.Authorization = new AuthenticationHeaderValue("Basic", Convert.ToBase64String(byteArray));
            hpcClient.DefaultRequestHeaders.Accept.Add(new MediaTypeWithQualityHeaderValue("application/json"));

            var stcContent = new StringContent(strJsonParameter, Encoding.UTF8, "application/json");
            hrmResult = await hpcClient.PostAsync(strDefFn + strFn, stcContent);

            return hrmResult;
        }


    }
}
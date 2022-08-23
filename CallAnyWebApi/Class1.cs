using System;
using System.Activities;
using System.Collections.Generic;
using System.Globalization;
using System.IO;
using System.Linq;
using System.Net;
using System.Text;
using System.Threading.Tasks;
using Microsoft.Xrm.Sdk;
using Microsoft.Xrm.Sdk.Query;
using Microsoft.Xrm.Sdk.Workflow;
using Newtonsoft.Json.Linq;
using RestSharp;
using RestSharp.Extensions;
using ApiLogger;

namespace NewCallAnyWebApi
{
    public class CallWebApiJsonSelectToken : CodeActivity
    {
        [RequiredArgument]
        [Input("Api")]
        [ReferenceTarget("chr_apiconfiguration")]
        public InArgument<EntityReference> Api { get; set; }

        [RequiredArgument]
        [Input("IsJSON")]
        [Default("true")]
        public InArgument<bool> IsJson { get; set; } = (InArgument<bool>)true;

        [Input("Body")]
        public InArgument<string> Body { get; set; }

        [Input("Find Mode")]
        [Default("true")]
        public InArgument<bool> FindMode { get; set; } = (InArgument<bool>)false;

        [Input("JSON Path")]
        [Default("succeed")]
        public InArgument<string> JsonPath { get; set; } 

        [Output("Response Content")]
        public OutArgument<string> ResponseContent { get; set; }

        [Output("JSON Path Value")]
        public OutArgument<string> JsonPathValue { get; set; }

        [Output("Result Status Code")]
        public OutArgument<string> ResultStatusCode { get; set; }

        protected override void Execute(CodeActivityContext executionContext)
        {
            IExecutionContext extension = executionContext.GetExtension<IExecutionContext>();
            IOrganizationServiceFactory extension2 = executionContext.GetExtension<IOrganizationServiceFactory>();
            IOrganizationService organizationService = extension2.CreateOrganizationService(extension.UserId);
            ITracingService extension3 = executionContext.GetExtension<ITracingService>();
            try
            {
                EntityReference entRefApi = Api.Get(executionContext);
                Entity entApiRetrieve = organizationService.Retrieve("chr_apiconfiguration", entRefApi.Id, new ColumnSet(true));


                string endPoint = entApiRetrieve["new_url"].ToString().Trim();
                string apiPath = entApiRetrieve["chr_name"].ToString().Trim();
                string methodType = entApiRetrieve.FormattedValues["new_verb"].ToUpper();

                string key1 = entApiRetrieve.Contains("new_key1") ? entApiRetrieve["new_key1"].ToString().Trim() : "X";
                string key2 = entApiRetrieve.Contains("new_key2") ? entApiRetrieve["new_key2"].ToString().Trim() : "X";
                string key3 = entApiRetrieve.Contains("new_key3") ? entApiRetrieve["new_key3"].ToString().Trim() : "X";
                string key4 = entApiRetrieve.Contains("new_key4") ? entApiRetrieve["new_key4"].ToString().Trim() : "X";
                string value1 = entApiRetrieve.Contains("new_value1") ? entApiRetrieve["new_value1"].ToString().Trim() : "X";
                string value2 = entApiRetrieve.Contains("new_value2") ? entApiRetrieve["new_value2"].ToString().Trim() : "X";
                string value3 = entApiRetrieve.Contains("new_value3") ? entApiRetrieve["new_value3"].ToString().Trim() : "X";
                string value4 = entApiRetrieve.Contains("new_value4") ? entApiRetrieve["new_value4"].ToString().Trim() : "X";

                string body = this.Body.Get<string>((ActivityContext)executionContext);
                string bodyForLog = body;
                bool isBodyJson = this.IsJson.Get<bool>((ActivityContext)executionContext);
                bool flag = this.FindMode.Get<bool>((ActivityContext)executionContext);
                string jsonPath = this.JsonPath.Get<string>((ActivityContext)executionContext);

                if (endPoint[endPoint.Length - 1] != '/')
                    endPoint = endPoint + '/';
                if (apiPath[0] == '/')
                    apiPath = apiPath.Substring(1);

                if (!isBodyJson) //if api is querystring 
                {
                    if (body != null)
                        endPoint = endPoint + apiPath + '?' + body;
                    else
                        endPoint = endPoint + apiPath;

                    body = null;
                }
                else
                {
                    endPoint = endPoint + apiPath;
                }

                Dictionary<string, string> headers = new Dictionary<string, string>();


                if (!string.IsNullOrEmpty(key1))
                    headers.Add(key1, value1);
                if (!string.IsNullOrEmpty(key2))
                    headers.Add(key2, value2);
                if (!string.IsNullOrEmpty(key3))
                    headers.Add(key3, value3);
                if (!string.IsNullOrEmpty(key4))
                    headers.Add(key4, value4);

                //Log
                //StreamWriter streamWriter = new StreamWriter("c:\\Temp\\API.log", true);
                //streamWriter.WriteLine("-----------------------");
                //streamWriter.WriteLine(endPoint);
                //streamWriter.WriteLine(body);
                //streamWriter.WriteLine(methodType);
                //streamWriter.WriteLine(headers[key1]);
                //streamWriter.WriteLine(headers[key2]);
                //streamWriter.WriteLine(headers[key3]);
                //streamWriter.WriteLine(headers[key4]);
                //streamWriter.Close();


                var strApiResult = CallApiMethod(endPoint, methodType, body, headers, isBodyJson);
                
                string strJsonPathValue = "";
                if (flag)
                {
                    List<JToken> tokens = this.FindTokens((JToken)JObject.Parse(strApiResult), jsonPath);
                    if (tokens.Count > 0)
                        strJsonPathValue = tokens[0].ToString();
                }
                else
                    strJsonPathValue = this.GetJsonSelectedToken(strApiResult, jsonPath);

                var stCode=GetJsonSelectedToken(strApiResult, "statusCode");
                
                //
                var logger = new LogAction();
                logger.SourceIp = Dns.GetHostName();
                logger.verb = methodType;
                logger.Url = entApiRetrieve["new_url"].ToString().Trim();
                logger.Path = apiPath;
                logger.RequestBody = bodyForLog;
                logger.ResponseBody = strApiResult.ToString();
                logger.Status = Int32.Parse(stCode);
                //(bool,string) suc=logger.AddLog(logger);
                logger.AddLog(logger);
                
                this.ResponseContent.Set((ActivityContext)executionContext, strApiResult);
                this.JsonPathValue.Set((ActivityContext)executionContext, strJsonPathValue);
                this.ResultStatusCode.Set((ActivityContext)executionContext, stCode);
                
            }
            catch (Exception ex)
            {
                throw ex;
            }
        }

        public string GetJsonSelectedToken(string jSonString, string jSonPath) => JObject.Parse(jSonString).SelectToken(jSonPath).Value<string>();

        public string CallApiMethod(string endPoint, string methodType, string body, Dictionary<string, string> headers, bool isJsonBody)
        {
            Method method = Method.POST;
            switch (methodType.ToUpper())
            {
                case "COPY":
                    //method = Method.COPY;
                    break;
                case "DELETE":
                    method = Method.DELETE;
                    break;
                case "GET":
                    method = Method.GET;
                    break;
                case "HEAD":
                    method = Method.HEAD;
                    break;
                case "MERGE":
                    //method = Method.Merge;
                    break;
                case "OPTIONS":
                    method = Method.OPTIONS;
                    break;
                case "PATCH":
                    method = Method.PATCH;
                    break;
                case "POST":
                    method = Method.POST;
                    break;
                case "PUT":
                    method = Method.PUT;
                    break;
            }

            var client = new RestClient(endPoint);
            var request = new RestRequest();

            client.Timeout = -1;
            request.Method = method;
            
            foreach (string str in headers.Keys.AsEnumerable<string>())
                request.AddHeader(str, headers[str]);
            
            if (!isJsonBody)
            {
                var response = client.Get(request);
                return response.Content;
            }
            else
            {
                request.AddParameter("application/json", body, ParameterType.RequestBody);
                var response = client.Execute((IRestRequest)request);
                return response.Content;
            }

        }

        public List<JToken> FindTokens(JToken containerToken, string name)
        {
            List<JToken> matches = new List<JToken>();
            this.FindTokens(containerToken, name, matches);
            return matches;
        }

        private void FindTokens(JToken containerToken, string name, List<JToken> matches)
        {
            if (containerToken.Type == JTokenType.Object)
            {
                foreach (JProperty child in containerToken.Children<JProperty>())
                {
                    if (child.Name == name)
                        matches.Add(child.Value);
                    this.FindTokens(child.Value, name, matches);
                }
            }
            else
            {
                if (containerToken.Type != JTokenType.Array)
                    return;
                foreach (JToken child in containerToken.Children())
                    this.FindTokens(child, name, matches);
            }
        }
    }
};
using Newtonsoft.Json;
using Newtonsoft.Json.Linq;
using System;
using System.Collections.Generic;
using System.Configuration;
using System.IO;
using System.Linq;
using System.Net;
using System.Net.Http;
using System.Text;

namespace SharepointTransfer
{
    public class Sharepoint 
    {
        public class SharepointParam
        {
            //public string _TYPE_ID;
            public string sharepointaaccessurl { get; set; }
            public string tenantid { get; set; }
            public string grant_type { get; set; }
            public string client_id { get; set; }
            public string client_secret { get; set; }
            public string resource { get; set; }
            public string access_token { get; set; }
            public string expires_in { get; set; }
            public string token_type { get; set; }
            public string expires_on { get; set; }
            public string sitedomain { get; set; }
            public string ServerRelativeUrl { get; set; }
            public string loginurl { get; set; }
            public string filepath { get; set; }
            public string filename { get; set; }
            public string RelativeAttchmntPath { get; set; }
            public string RelativeAttchmntPathDestination { get; set; }
            public string BookingListId { get; set; }
            public string client_id_destination { get; set; }
            public string client_secret_destination { get; set; }
            public string destinationsharepointurl { get; set; }
            public string resource_destination { get; set; }
            public string sitedomain_destination { get; set; }
            public string tenantid_destination { get; set; }
            public string ServerRelativeUrl_destination { get; set; }
            public string siteuri { get; set; }
            public string FileFolder { get; set; }
            public string SubFolder { get; set; }
            public string access_token_destination { get; set; }
            private string _Fin_Year;
            public string Fin_Year
            {
                get { return _Fin_Year; }
                set { _Fin_Year = value; }
            }

            private string _TYPE_ID { get; set; }
            public string TYPE_ID
            {
                get { return _TYPE_ID; }
                set { _TYPE_ID = value; }
            }
            private string _PROCESS_NUMBER { get; set; }
            public string PROCESS_NUMBER
            {
                get { return _PROCESS_NUMBER; }
                set { _PROCESS_NUMBER = value; }
            }
            private string _CREATED_BY { get; set; }
            public string CREATED_BY
            {
                get { return _CREATED_BY; }
                set { _CREATED_BY = value; }
            }
            private string _CREATED_DATE { get; set; }
            public string CREATED_DATE
            {
                get { return _CREATED_DATE; }
                set { _CREATED_DATE = value; }
            }
            private string _ENTITY_CODE { get; set; }
            public string ENTITY_CODE
            {
                get { return _ENTITY_CODE; }
                set { _ENTITY_CODE = value; }
            }
            private string _Entity { get; set; }
            public string Entity
            {
                get { return _Entity; }
                set { _Entity = value; }
            }
            private string _Entity_GST_Number { get; set; }
            public string Entity_GST_Number
            {
                get { return _Entity_GST_Number; }
                set { _Entity_GST_Number = value; }
            }
            private string _Function { get; set; }
            public string Function
            {
                get { return _Function; }
                set { _Function = value; }
            }
            private string _Invoice_Amount { get; set; }
            public string Invoice_Amount
            {
                get { return _Invoice_Amount; }
                set { _Invoice_Amount = value; }
            }
            private string _Invoice_Date { get; set; }
            public string Invoice_Date
            {
                get { return _Invoice_Date; }
                set { _Invoice_Date = value; }
            }
            private string _Modified { get; set; }
            public string Modified
            {
                get { return _Modified; }
                set { _Modified = value; }
            }
            private string _Invoice_Number { get; set; }
            public string Invoice_Number
            {
                get { return _Invoice_Number; }
                set { _Invoice_Number = value; }
            }
            private string _Month { get; set; }
            public string Month
            {
                get { return _Month; }
                set { _Month = value; }
            }
            private string _PAN_Number { get; set; }
            public string PAN_Number
            {
                get { return _PAN_Number; }
                set { _PAN_Number = value; }
            }
            private string _Party_GST_Number { get; set; }
            public string Party_GST_Number
            {
                get { return _Party_GST_Number; }
                set { _Party_GST_Number = value; }
            }

            private string _Reference_Number_Name { get; set; }
            public string Reference_Number_Name
            {
                get { return _Reference_Number_Name; }
                set { _Reference_Number_Name = value; }
            }
            private string _Sub_Function { get; set; }
            public string Sub_Function
            {
                get { return _Sub_Function; }
                set { _Sub_Function = value; }
            }
            private string _Vendor_Name { get; set; }
            public string Vendor_Name
            {
                get { return _Vendor_Name; }
                set { _Vendor_Name = value; }
            }
            private string _AP_DMS_Created { get; set; }
            public string AP_DMS_Created
            {
                get { return _AP_DMS_Created; }
                set { _AP_DMS_Created = value; }
            }
            private string _AP_DMS_Modified_Date { get; set; }
            public string AP_DMS_Modified_Date
            {
                get { return _AP_DMS_Modified_Date; }
                set { _AP_DMS_Modified_Date = value; }
            }
            private string _AP_DMS_CreatedDate { get; set; }
            public string AP_DMS_CreatedDate
            {
                get { return _AP_DMS_CreatedDate; }
                set { _AP_DMS_CreatedDate = value; }
            }
            private string _AP_DMS_LASTMODIFIER { get; set; }
            public string AP_DMS_LASTMODIFIER
            {
                get { return _AP_DMS_LASTMODIFIER; }
                set { _AP_DMS_LASTMODIFIER = value; }
            }
            private string _Reference_No { get; set; }
            public string Reference_No
            {
                get { return _Reference_No; }
                set { _Reference_No = value; }
            }
            private string _TDS_Amount { get; set; }
            public string TDS_Amount
            {
                get { return _TDS_Amount; }
                set { _TDS_Amount = value; }
            }
            private string _State { get; set; }
            public string State
            {
                get { return _State; }
                set { _State = value; }
            }
            private string _ClassID { get; set; }
            public string ClassID
            {
                get { return _ClassID; }
                set { _ClassID = value; }
            }

            public string UpdatePropertyURL = string.Empty;
            public FileProperty ObjFileProperty { get; set; }
            public string ListItem_Metadata_Type = string.Empty;
        }


        public class entityDTO
        {
            public List<int> ID { get; set; }
            public EntityMain Entity { get; set; }
        }

        public class EntityMain
        {
            public string Entity { get; set; }
        }

        public class FileProperty
        {
            public FileProperty()
            {
                __metadata = new CreateFolderMetadata();
            }
            public CreateFolderMetadata __metadata { get; set; }

            private string _Fin_Year;
            public string Fin_Year
            {
                get { return _Fin_Year; }
                set { _Fin_Year = value; }
            }

            private string _TYPE_ID { get; set; }
            public string TYPE_ID
            {
                get { return _TYPE_ID; }
                set { _TYPE_ID = value; }
            }
            private string _PROCESS_NUMBER { get; set; }
            public string PROCESS_NUMBER
            {
                get { return _PROCESS_NUMBER; }
                set { _PROCESS_NUMBER = value; }
            }
            private string _CREATED_BY { get; set; }
            public string CREATED_BY
            {
                get { return _CREATED_BY; }
                set { _CREATED_BY = value; }
            }
            private string _CREATED_DATE { get; set; }
            public string CREATED_DATE
            {
                get { return _CREATED_DATE; }
                set { _CREATED_DATE = value; }
            }
            private string _ENTITY_CODE { get; set; }
            public string ENTITY_CODE
            {
                get { return _ENTITY_CODE; }
                set { _ENTITY_CODE = value; }
            }
            private string _Entity { get; set; }
            public string Entity
            {
                get { return _Entity; }
                set { _Entity = value; }
            }
            private string _Entity_GST_Number { get; set; }
            public string Entity_GST_Number
            {
                get { return _Entity_GST_Number; }
                set { _Entity_GST_Number = value; }
            }
            private string _Function { get; set; }
            public string Function
            {
                get { return _Function; }
                set { _Function = value; }
            }
            private string _Invoice_Amount { get; set; }
            public string Invoice_Amount
            {
                get { return _Invoice_Amount; }
                set { _Invoice_Amount = value; }
            }
            private string _Invoice_Date { get; set; }
            public string Invoice_Date
            {
                get { return _Invoice_Date; }
                set { _Invoice_Date = value; }
            }
            private string _Invoice_Number { get; set; }
            public string Invoice_Number
            {
                get { return _Invoice_Number; }
                set { _Invoice_Number = value; }
            }
            private string _Month { get; set; }
            public string Month
            {
                get { return _Month; }
                set { _Month = value; }
            }
            private string _PAN_Number { get; set; }
            public string PAN_Number
            {
                get { return _PAN_Number; }
                set { _PAN_Number = value; }
            }
            private string _Party_GST_Number { get; set; }
            public string Party_GST_Number
            {
                get { return _Party_GST_Number; }
                set { _Party_GST_Number = value; }
            }

            private string _Reference_Number_Name { get; set; }
            public string Reference_Number_Name
            {
                get { return _Reference_Number_Name; }
                set { _Reference_Number_Name = value; }
            }
            private string _Sub_Function { get; set; }
            public string Sub_Function
            {
                get { return _Sub_Function; }
                set { _Sub_Function = value; }
            }
            private string _Vendor_Name { get; set; }
            public string Vendor_Name
            {
                get { return _Vendor_Name; }
                set { _Vendor_Name = value; }
            }
            private string _AP_DMS_Created { get; set; }
            public string AP_DMS_Created
            {
                get { return _AP_DMS_Created; }
                set { _AP_DMS_Created = value; }
            }
            private string _AP_DMS_Modified_Date { get; set; }
            public string AP_DMS_Modified_Date
            {
                get { return _AP_DMS_Modified_Date; }
                set { _AP_DMS_Modified_Date = value; }
            }
            private string _AP_DMS_CreatedDate { get; set; }
            public string AP_DMS_CreatedDate
            {
                get { return _AP_DMS_CreatedDate; }
                set { _AP_DMS_CreatedDate = value; }
            }
            private string _AP_DMS_LASTMODIFIER { get; set; }
            public string AP_DMS_LASTMODIFIER
            {
                get { return _AP_DMS_LASTMODIFIER; }
                set { _AP_DMS_LASTMODIFIER = value; }
            }
            private string _Reference_No { get; set; }
            public string Reference_No
            {
                get { return _Reference_No; }
                set { _Reference_No = value; }
            }
            private string _TDS_Amount { get; set; }
            public string TDS_Amount
            {
                get { return _TDS_Amount; }
                set { _TDS_Amount = value; }
            }
            private string _State { get; set; }
            public string State
            {
                get { return _State; }
                set { _State = value; }
            }
            private string _ClassID { get; set; }
            public string ClassID
            {
                get { return _ClassID; }
                set { _ClassID = value; }
            }


        }

        public class CreateFolderMetadata
        {
            public string type { get; set; }
        }

        public string SharepointCreateFolderForFilesOnly(SharepointParam objSharepointParam)
        {
            try
            {
                string FolderFullPath = objSharepointParam.RelativeAttchmntPathDestination;
                FolderFullPath = FolderFullPath.Replace("\\", "/").Replace("//", "/");
                SharepointCreateInternalFolderForSubFolder(objSharepointParam, FolderFullPath, "");
                SharepointCreateInternalFolderForFilesOnly(objSharepointParam, FolderFullPath, "");
                return "";
            }
            catch (Exception ex)
            {

                throw ex;

            }
        }

        public string SharepointCreateInternalFolderForSubFolder(SharepointParam objSharepointParam, string FullFolderPath, string ParentFolder)
        {
            try
            {
                var folderUrls = FullFolderPath.Split(new char[] { '/' }, StringSplitOptions.RemoveEmptyEntries);

                SharepointParam objparam = objSharepointParam;
                objSharepointParam.ServerRelativeUrl_destination = objSharepointParam.ServerRelativeUrl_destination.TrimEnd(new char[] { '/' });
                objparam.ServerRelativeUrl_destination = objSharepointParam.ServerRelativeUrl_destination + "/" + folderUrls[0];

                JToken __token = new JObject();
                JToken __tokenType = new JObject();
                __tokenType["type"] = "SP.Folder";
                __token["ServerRelativeUrl"] = objSharepointParam.ServerRelativeUrl_destination;//"/sites/Exponentia-DMS/Exponentia/Installation_10";
                __token["__metadata"] = __tokenType;
                var jsonString = Newtonsoft.Json.JsonConvert.SerializeObject(__token);
                APICallSharepoint objApiCall = new APICallSharepoint();
                List<APIParameters> lstParameter = new List<APIParameters>();
                APIParameters objApiParameter = new APIParameters();
                objApiParameter.PropertyName = "Authorization";
                objApiParameter.PropertyValue = "Bearer " + objSharepointParam.access_token_destination;
                lstParameter.Add(objApiParameter);

                objApiParameter = new APIParameters();
                objApiParameter.PropertyName = "X-RequestDigest";
                objApiParameter.PropertyValue = "formDigest";
                lstParameter.Add(objApiParameter);

                byte[] data = Encoding.ASCII.GetBytes(jsonString);
                string response = objApiCall.httpwebrequest(objSharepointParam.destinationsharepointurl + "_api/web/folders", lstParameter, "application/json;odata=verbose", "application/json;odata=verbose", "Post", data);


                return response;

            }
            catch (Exception ex)
            {

                throw ex;
            }
        }


        public string SharepointCreateInternalFolderForFilesOnly(SharepointParam objSharepointParam, string FullFolderPath, string ParentFolder)
        {
            try
            {
                var folderUrls = FullFolderPath.Split(new char[] { '/' }, StringSplitOptions.RemoveEmptyEntries);

                SharepointParam objparam = objSharepointParam;
                objSharepointParam.ServerRelativeUrl_destination = objSharepointParam.ServerRelativeUrl_destination.TrimEnd(new char[] { '/' });
                objparam.ServerRelativeUrl_destination = objSharepointParam.ServerRelativeUrl_destination + "/" + folderUrls[1];

                JToken __token = new JObject();
                JToken __tokenType = new JObject();
                __tokenType["type"] = "SP.Folder";
                __token["ServerRelativeUrl"] = objSharepointParam.ServerRelativeUrl_destination;//"/sites/Exponentia-DMS/Exponentia/Installation_10";
                __token["__metadata"] = __tokenType;
                var jsonString = Newtonsoft.Json.JsonConvert.SerializeObject(__token);
                APICallSharepoint objApiCall = new APICallSharepoint();
                List<APIParameters> lstParameter = new List<APIParameters>();
                APIParameters objApiParameter = new APIParameters();
                objApiParameter.PropertyName = "Authorization";
                objApiParameter.PropertyValue = "Bearer " + objSharepointParam.access_token_destination;
                lstParameter.Add(objApiParameter);

                objApiParameter = new APIParameters();
                objApiParameter.PropertyName = "X-RequestDigest";
                objApiParameter.PropertyValue = "formDigest";
                lstParameter.Add(objApiParameter);

                byte[] data = Encoding.ASCII.GetBytes(jsonString);
                string response = objApiCall.httpwebrequest(objSharepointParam.destinationsharepointurl + "_api/web/folders", lstParameter, "application/json;odata=verbose", "application/json;odata=verbose", "Post", data);

                return response;

            }
            catch (Exception ex)
            {

                throw ex;
            }
        }


        public class SharepointData
        {
            public string id { get; set; }
        }
        
        public class SharepointList
        {
            public string text { get; set; }
            public List<SharepointValue> value { get; set; }
        }

        public class ListData
        {
            public string text { get; set; }
            public List<ListDataValue> value { get; set; }
        }

        public class ListDataValue
        {
            public FieldListData fields { get; set; }

        }

        public class FieldListData
        {
            public string FileLeafRef { get; set; }
            public string id { get; set; }
        }

        public class SharepointValue
        {
            public string id { get; set; }
            public string displayName { get; set; }
        }

        public SharepointParam CreateSharepointObject()
        {
            Console.WriteLine("CreateSharePointObject Started");
            SharepointParam obj = new SharepointParam();
            obj.client_id = ConfigurationManager.AppSettings["CLIENT_ID"];
            obj.client_secret = ConfigurationManager.AppSettings["CLIENT_SECRET"];
            obj.resource = ConfigurationManager.AppSettings["RESOURCE"];
            obj.sitedomain = ConfigurationManager.AppSettings["SITEDOMAIN"];
            obj.tenantid = ConfigurationManager.AppSettings["TENANTID"];
            obj.ServerRelativeUrl = ConfigurationManager.AppSettings["LEADLIBRARYNAME"].ToString();
            obj.sharepointaaccessurl = ConfigurationManager.AppSettings["SharePointURL"].ToString();
            obj.destinationsharepointurl = ConfigurationManager.AppSettings["DestinationSharePointURL"].ToString();
            obj.client_id_destination = ConfigurationManager.AppSettings["CLIENT_ID_DESTINATION"];
            obj.client_secret_destination = ConfigurationManager.AppSettings["CLIENT_SECRET_DESTINATION"];
            obj.resource_destination = ConfigurationManager.AppSettings["RESOURCE_DESTINATION"];
            obj.sitedomain_destination = ConfigurationManager.AppSettings["SITEDOMAIN_DESTINATION"];
            obj.tenantid_destination = ConfigurationManager.AppSettings["TENANTID_DESTINATION"];
            obj.ServerRelativeUrl_destination = ConfigurationManager.AppSettings["LEADLIBRARYNAME_DESTINATION"].ToString();
            obj.siteuri = ConfigurationManager.AppSettings["BASEURI"].ToString();
            Console.WriteLine("CreateSharePointObject Ended");
            return obj;
        }



        public string SharepointAccess(SharepointParam objSharepointParam)

        {
            string response = string.Empty;
            try
            {
                ServicePointManager.SecurityProtocol = (SecurityProtocolType)3072;
                WebRequest myWebRequest;

                string stGetAccessTokenUrl = "https://accounts.accesscontrol.windows.net/{{account_id}}/tokens/OAuth/2";
                myWebRequest = WebRequest.Create(stGetAccessTokenUrl);
                myWebRequest.ContentType = "application/x-www-form-urlencoded";
                myWebRequest.Method = "POST";
                myWebRequest.Timeout = 300000;
                var postData = "grant_type=client_credentials";
                postData += "&client_id=" + objSharepointParam.client_id + "@" + objSharepointParam.tenantid;
                postData += "&client_secret=" + objSharepointParam.client_secret;
                postData += "&resource=" + objSharepointParam.resource + "/" + objSharepointParam.sitedomain + "@" + objSharepointParam.tenantid;

                var data = Encoding.ASCII.GetBytes(postData);
                using (var stream = myWebRequest.GetRequestStream())
                {
                    stream.Write(data, 0, data.Length);
                }

                HttpWebResponse endpointResponse = (HttpWebResponse)myWebRequest.GetResponse();
                response = new StreamReader(endpointResponse.GetResponseStream()).ReadToEnd();
            }
            catch (Exception ex)
            {
                Console.WriteLine("Error In RetrieveAccessCode:" + ex.Message + "-" + ex.StackTrace, "");
            }
            return response;

        }

        public string SharepointAccessForDestination(SharepointParam objSharepointParam)
        {
            string response = string.Empty;
            try
            {
                ServicePointManager.SecurityProtocol = (SecurityProtocolType)3072;
                WebRequest myWebRequest1;
                string stGetAccessTokenUrl1 = "https://accounts.accesscontrol.windows.net/{{account_id}}/tokens/OAuth/2"; ;
                myWebRequest1 = WebRequest.Create(stGetAccessTokenUrl1);
                myWebRequest1.ContentType = "application/x-www-form-urlencoded";
                myWebRequest1.Method = "POST";
                myWebRequest1.Timeout = 300000;
                var postData1 = "grant_type=client_credentials";

                postData1 += "&client_id=" + objSharepointParam.client_id_destination + "@" + objSharepointParam.tenantid_destination;
                postData1 += "&client_secret=" + objSharepointParam.client_secret_destination;
                postData1 += "&resource=" + objSharepointParam.resource_destination + "/" + objSharepointParam.sitedomain_destination + "@" + objSharepointParam.tenantid_destination;
                var data1 = Encoding.ASCII.GetBytes(postData1);
                using (var stream = myWebRequest1.GetRequestStream())
                {
                    stream.Write(data1, 0, data1.Length);
                }

                HttpWebResponse endpointResponse = (HttpWebResponse)myWebRequest1.GetResponse();
                response = new StreamReader(endpointResponse.GetResponseStream()).ReadToEnd();
            }
            catch (Exception ex)
            {
                Console.WriteLine("Error In RetrieveAccessCode:" + ex.Message + "-" + ex.StackTrace, "");
            }
            return response;

        }

        public string SharepointFolderExistsForDestination(SharepointParam objSharepointParam)
        {
            try
            {
                string FolderPath = objSharepointParam.RelativeAttchmntPathDestination;
                FolderPath = FolderPath.Replace("\\", "/").Replace("//", "/");
                string[] folderName = FolderPath.Split("/"); 
                APICallSharepoint objApiCall = new APICallSharepoint();
                List<APIParameters> lstParameter = new List<APIParameters>();
                APIParameters objApiParameter = new APIParameters();
                objApiParameter.PropertyName = "Authorization";
                objApiParameter.PropertyValue = "Bearer " + objSharepointParam.access_token_destination;
                lstParameter.Add(objApiParameter);
                HttpWebRequest endpointRequestForList = (HttpWebRequest)HttpWebRequest.Create(objSharepointParam.destinationsharepointurl + "_api/Web/GetFolderByServerRelativePath(decodedurl='" + objSharepointParam.ServerRelativeUrl_destination + folderName[0] + "')/Exists");
                endpointRequestForList.Method = "GET";
                endpointRequestForList.Accept = "application/json;odata=verbose";
                endpointRequestForList.Headers.Add("Authorization", "Bearer " + objSharepointParam.access_token_destination);
                HttpWebResponse endpointResponseForList = (HttpWebResponse)endpointRequestForList.GetResponse();
                Stream webStreamList = endpointResponseForList.GetResponseStream();
                StreamReader responseReaderList = new StreamReader(webStreamList);
                string responseFileList = responseReaderList.ReadToEnd();
                JObject jobjList = JObject.Parse(responseFileList);
                JArray jarr = (JArray)jobjList["d"]["results"];
                string response = objApiCall.httpwebrequest(objSharepointParam.destinationsharepointurl + "_api/web/GetFolderByServerRelativeUrl('" + objSharepointParam.ServerRelativeUrl_destination + FolderPath + "/')", lstParameter, "", "application/json;odata=verbose", "Get", null);

                return response;
            }
            catch (Exception ex)
            {

                return "";
            }
        }

        public class APICallSharepoint
        {
            public string httpwebrequest(string Url, List<APIParameters> lstHeader, string ContentType, string Accept, string Method, byte[] _ByteArray)
            
            {
                try
                {
                    SetSSL();
                    ServicePointManager.SecurityProtocol = SecurityProtocolType.Tls | SecurityProtocolType.Tls11 | SecurityProtocolType.Tls12;
                    HttpWebRequest httpWReq = (HttpWebRequest)WebRequest.Create(Url);
                    Encoding encoding = new UTF8Encoding();
                    httpWReq.ProtocolVersion = HttpVersion.Version11;
                    httpWReq.Method = Method;
                    if (Accept != "")
                    {
                        httpWReq.Accept = Accept;
                    }
                    if (ContentType != "")
                    {
                        httpWReq.ContentType = ContentType;
                    }
                    WebAddHeader(httpWReq, lstHeader);
                    if (Method == "Post")
                    {
                        httpWReq.ContentLength = _ByteArray.Length;
                        Stream requestStream = httpWReq.GetRequestStream();
                        requestStream.Write(_ByteArray, 0, _ByteArray.Length);
                        requestStream.Close();
                    }
                    ServicePointManager.SecurityProtocol = SecurityProtocolType.Tls | SecurityProtocolType.Tls11 | SecurityProtocolType.Tls12;
                    HttpWebResponse response = (HttpWebResponse)httpWReq.GetResponse();
                    var responseString = new StreamReader(response.GetResponseStream()).ReadToEnd();
                    return responseString;
                }
                catch (Exception ex)
                {
                    return "";
                }
            }

            private void SetSSL()
            {
                ServicePointManager.SecurityProtocol = (SecurityProtocolType)3072;
            }

            private static void WebAddHeader(HttpWebRequest client, List<APIParameters> lstHeader)
            {
                if (lstHeader != null && lstHeader.Count > 0)
                {
                    foreach (APIParameters par in lstHeader)
                    {
                        client.Headers.Add(par.PropertyName, par.PropertyValue);
                    }
                }
            }

        }

        public string setEntities(SharepointParam objSharepointParam, int entID)
        {
            try
            {
                int ent_ID = 0;
                string ent_Name = "";
                ServicePointManager.SecurityProtocol = SecurityProtocolType.Tls | SecurityProtocolType.Tls11 | SecurityProtocolType.Tls12;
                HttpWebRequest endpointRequestForList = (HttpWebRequest)HttpWebRequest.Create(objSharepointParam.sharepointaaccessurl + "_api/web/lists");
                endpointRequestForList.Method = "GET";
                endpointRequestForList.Accept = "application/json;odata=verbose";
                endpointRequestForList.Headers.Add("Authorization", "Bearer " + objSharepointParam.access_token);
                HttpWebResponse endpointResponseForList = (HttpWebResponse)endpointRequestForList.GetResponse();
                Stream webStreamList = endpointResponseForList.GetResponseStream();
                StreamReader responseReaderList = new StreamReader(webStreamList);
                string responseFileList = responseReaderList.ReadToEnd();
                JObject jobjList = JObject.Parse(responseFileList);

                JArray jarr = (JArray)jobjList["d"]["results"];
                foreach (JObject j in jarr)
                {
                    string entityList = (string)j["Title"];
                    if (entityList == "Entity")
                    {
                        ServicePointManager.SecurityProtocol = SecurityProtocolType.Tls | SecurityProtocolType.Tls11 | SecurityProtocolType.Tls12;
                        string itemEntityURL = (string)j["Items"]["__deferred"]["uri"];
                        HttpWebRequest entityItemRequest = (HttpWebRequest)HttpWebRequest.Create(itemEntityURL);
                        entityItemRequest.Method = "GET";
                        entityItemRequest.Accept = "application/json;odata=verbose";
                        entityItemRequest.Headers.Add("Authorization", "Bearer " + objSharepointParam.access_token);
                        HttpWebResponse entityItemResponse = (HttpWebResponse)entityItemRequest.GetResponse();
                        Stream webStreamRes = entityItemResponse.GetResponseStream();
                        StreamReader responseReader = new StreamReader(webStreamRes);
                        string responseFile = responseReader.ReadToEnd();
                        JObject jobjEntity = JObject.Parse(responseFile);
                        JArray jarrList = (JArray)jobjEntity["d"]["results"];
                        foreach (JObject k in jarrList)
                        {
                            ent_ID = (int)k["ID"];
                            ent_Name = (string)k["Entity"];
                            if (ent_ID == entID)
                                return ent_Name;
                        }
                    }
                }
                return ent_Name;
            }
            catch(Exception ex)
            {
                Console.WriteLine("Exception in Entity set " + ex.Message);
                return null;   
            }
            
        }

        public void getSharepointFilesOnly(SharepointParam objSharepointParam)
        {
            try
            {
                SharepointData data = new SharepointData();
                SharepointList sharepointList = new SharepointList();
                SharepointValue shareValue = new SharepointValue();

                #region To get Data(docs,pdf,xls,etc) from List menu
                ServicePointManager.SecurityProtocol = SecurityProtocolType.Tls | SecurityProtocolType.Tls11 | SecurityProtocolType.Tls12;
                Console.WriteLine("getByteDataFromList Started");
                HttpWebRequest endpointRequestForList = (HttpWebRequest)HttpWebRequest.Create(objSharepointParam.sharepointaaccessurl + "_api/Web/GetFolderByServerRelativePath(decodedurl='" + objSharepointParam.ServerRelativeUrl + "')/Folders");
                endpointRequestForList.Method = "GET";
                endpointRequestForList.Accept = "application/json;odata=verbose";
                endpointRequestForList.Headers.Add("Authorization", "Bearer " + objSharepointParam.access_token);
                HttpWebResponse endpointResponseForList = (HttpWebResponse)endpointRequestForList.GetResponse();
                Stream webStreamList = endpointResponseForList.GetResponseStream();
                StreamReader responseReaderList = new StreamReader(webStreamList);
                string responseFileList = responseReaderList.ReadToEnd();
                JObject jobjList = JObject.Parse(responseFileList);
                JArray jarr = (JArray)jobjList["d"]["results"];
                foreach (JObject a in jarr)
                {
                    //string[] ListArrName = { "201920", "APDMSLIB", "APSDMSLIB", "APDMSOLDPARTI", "DMSLIBA", "DMSLIBB", "DMSLIBC", "DMSLIBD", "DMSLIBE", "DMSLIBF", "DMSLIBH", "DMSLIBJ", "Scaned_Metadata" };
                    string ListName = (string)a["Name"];

                    string[] ListArrName = {"DMSLIBD", "DMSLIBE" };

                    if (ListArrName.Contains(ListName))
                    {
                        string ServerRelativeUrl = "";
                        string ListFolderURL = (string)a["Folders"]["__deferred"]["uri"];
                        ServicePointManager.SecurityProtocol = SecurityProtocolType.Tls | SecurityProtocolType.Tls11 | SecurityProtocolType.Tls12;
                        HttpWebRequest endpointRequestForListTitle = (HttpWebRequest)HttpWebRequest.Create(objSharepointParam.sharepointaaccessurl + "_api/Web/GetFolderByServerRelativePath(decodedurl='" + objSharepointParam.ServerRelativeUrl + ListName.ToString() + "')/Properties");
                        endpointRequestForListTitle.Method = "GET";
                        endpointRequestForListTitle.Accept = "application/json;odata=verbose";
                        endpointRequestForListTitle.Headers.Add("Authorization", "Bearer " + objSharepointParam.access_token);
                        HttpWebResponse endpointResponseForListTitle = (HttpWebResponse)endpointRequestForListTitle.GetResponse();
                        Stream webStreamForListTitle = endpointResponseForListTitle.GetResponseStream();
                        StreamReader responseReaderForListTitle = new StreamReader(webStreamForListTitle);
                        string responseForListTitle = responseReaderForListTitle.ReadToEnd();
                        JObject jobjForListTitle = JObject.Parse(responseForListTitle);
                        string listName = (string)jobjForListTitle["d"]["vti_x005f_listtitle"];

                        ServicePointManager.SecurityProtocol = SecurityProtocolType.Tls | SecurityProtocolType.Tls11 | SecurityProtocolType.Tls12;
                        HttpWebRequest folderListDetails = (HttpWebRequest)HttpWebRequest.Create(ListFolderURL);
                        folderListDetails.Method = "GET";
                        folderListDetails.Accept = "application/json;odata=verbose";
                        folderListDetails.Headers.Add("Authorization", "Bearer " + objSharepointParam.access_token);
                        HttpWebResponse endpointfolderListDetails = (HttpWebResponse)folderListDetails.GetResponse();
                        Stream streamFolderListDetails = endpointfolderListDetails.GetResponseStream();
                        StreamReader responseFolderListDetails = new StreamReader(streamFolderListDetails);
                        string resFolderListDetails = responseFolderListDetails.ReadToEnd();
                        JObject jobjFolderListDetails = JObject.Parse(resFolderListDetails);
                        JArray jarrList = (JArray)jobjFolderListDetails["d"]["results"];
                        foreach (JObject j in jarrList)
                        {
                            string uriList = (string)j["Files"]["__deferred"]["uri"];
                            string folderName = (string)j["Name"];

                            if (folderName == "Future" || folderName == "Taxation")
                            {
                                Console.WriteLine("FolderName " + folderName);
                                ServicePointManager.SecurityProtocol = SecurityProtocolType.Tls | SecurityProtocolType.Tls11 | SecurityProtocolType.Tls12;
                                HttpWebRequest endpointRequestForFoldersItems = (HttpWebRequest)HttpWebRequest.Create(uriList);
                                endpointRequestForFoldersItems.Method = "GET";
                                endpointRequestForFoldersItems.Accept = "application/json;odata=verbose";
                                endpointRequestForFoldersItems.Headers.Add("Authorization", "Bearer " + objSharepointParam.access_token);
                                HttpWebResponse endpointResponseForFolderItems = (HttpWebResponse)endpointRequestForFoldersItems.GetResponse();
                                Stream webStreamFolderItems = endpointResponseForFolderItems.GetResponseStream();
                                StreamReader responseReaderFolderItems = new StreamReader(webStreamFolderItems);
                                string responseFileFolderItems = responseReaderFolderItems.ReadToEnd();
                                JObject jobjFolderItems = JObject.Parse(responseFileFolderItems);
                                JArray jarrFolder = (JArray)jobjFolderItems["d"]["results"];
                                foreach (JObject l in jarrFolder)
                                {
                                    string uri = (string)l["__metadata"]["id"];
                                    string ListItemsMetaData = (string)l["ListItemAllFields"]["__deferred"]["uri"];
                                    string fileName = (string)l["Name"];
                                    Console.WriteLine("Filename " + fileName);
                                    ServicePointManager.SecurityProtocol = SecurityProtocolType.Tls | SecurityProtocolType.Tls11 | SecurityProtocolType.Tls12;
                                    HttpWebRequest endpointRequestForProperties = (HttpWebRequest)HttpWebRequest.Create(uri + "/ListItemAllFields");
                                    endpointRequestForProperties.Method = "GET";
                                    endpointRequestForProperties.Accept = "application/json;odata=verbose";
                                    endpointRequestForProperties.Headers.Add("Authorization", "Bearer " + objSharepointParam.access_token);
                                    HttpWebResponse endpointResponseForProperties = (HttpWebResponse)endpointRequestForProperties.GetResponse();
                                    Stream webStreamForProperties = endpointResponseForProperties.GetResponseStream();
                                    StreamReader responseReaderForProperties = new StreamReader(webStreamForProperties);
                                    string responseFileForProperties = responseReaderForProperties.ReadToEnd();
                                    JObject jobjForProperties = JObject.Parse(responseFileForProperties);
                                    string entityName = "";
                                    if ((int)jobjForProperties["d"]["EntityId"] == null || (int)jobjForProperties["d"]["EntityId"] == 0)
                                    {

                                    }
                                    else
                                    {
                                        int entityID = 0;
                                        entityID = (int)jobjForProperties["d"]["EntityId"];
                                        objSharepointParam.Entity = (string)jobjForProperties["d"]["EntityName"];
                                        objSharepointParam.Entity = setEntities(objSharepointParam, entityID);
                                        entityName = objSharepointParam.Entity;
                                    }
                                    if ((string)jobjForProperties["d"]["Entity_x0020_GST_x0020_Number"] == null)
                                    {

                                    }
                                    else
                                    {
                                        objSharepointParam.Entity_GST_Number = (string)jobjForProperties["d"]["Entity_x0020_GST_x0020_Number"];
                                    }
                                    if ((string)jobjForProperties["d"]["Modified"] == null)
                                    {

                                    }
                                    else
                                    {
                                        objSharepointParam.Modified = (string)jobjForProperties["d"]["Modified"];
                                    }
                                    if ((string)jobjForProperties["d"]["Fin_x0020_Year"] == null)
                                    {

                                    }
                                    else
                                    {
                                        objSharepointParam.Fin_Year = (string)jobjForProperties["d"]["Fin_x0020_Year"];
                                    }
                                    if ((string)jobjForProperties["d"]["Function"] == null)
                                    {

                                    }
                                    else
                                    {
                                        objSharepointParam.Function = (string)jobjForProperties["d"]["Function"];
                                    }
                                    if ((string)jobjForProperties["d"]["Invoice_x0020_Amount"] == null)
                                    {

                                    }
                                    else
                                    {
                                        objSharepointParam.Invoice_Amount = (string)jobjForProperties["d"]["Invoice_x0020_Amount"];
                                    }
                                    if ((string)jobjForProperties["d"]["Invoice_x0020_Number"] == null)
                                    {

                                    }
                                    else
                                    {
                                        objSharepointParam.Invoice_Number = (string)jobjForProperties["d"]["Invoice_x0020_Number"];
                                    }
                                    if ((string)jobjForProperties["d"]["Invoice_x0020_Date"] == null)
                                    {

                                    }
                                    else
                                    {
                                        objSharepointParam.Invoice_Date = (string)jobjForProperties["d"]["Invoice_x0020_Date"];
                                    }
                                    if ((string)jobjForProperties["d"]["Month"] == null)
                                    {

                                    }
                                    else
                                    {
                                        objSharepointParam.Month = (string)jobjForProperties["d"]["Month"];
                                    }
                                    if ((string)jobjForProperties["d"]["PAN_x0020_Number"] == null)
                                    {

                                    }
                                    else
                                    {
                                        objSharepointParam.PAN_Number = (string)jobjForProperties["d"]["PAN_x0020_Number"];
                                    }
                                    if ((string)jobjForProperties["d"]["Party_x0020_GST_x0020_Number"] == null)
                                    {

                                    }
                                    else
                                    {
                                        objSharepointParam.Party_GST_Number = (string)jobjForProperties["d"]["Party_x0020_GST_x0020_Number"];
                                    }
                                    if ((string)jobjForProperties["d"]["Reference_x0020_Number_Name"] == null)
                                    {

                                    }
                                    else
                                    {
                                        objSharepointParam.Reference_Number_Name = (string)jobjForProperties["d"]["Reference_x0020_Number_Name"];
                                    }
                                    if ((string)jobjForProperties["d"]["Sub_x0020_Function"] == null)
                                    {

                                    }
                                    else
                                    {
                                        objSharepointParam.Sub_Function = (string)jobjForProperties["d"]["Sub_x0020_Function"];
                                    }
                                    if ((string)jobjForProperties["d"]["Vendor_x0020_Name"] == null)
                                    {

                                    }
                                    else
                                    {
                                        objSharepointParam.Vendor_Name = (string)jobjForProperties["d"]["Vendor_x0020_Name"];
                                    }
                                    if ((string)jobjForProperties["d"]["AP_DMS_Created"] == null)
                                    {

                                    }
                                    else
                                    {
                                        objSharepointParam.AP_DMS_Created = (string)jobjForProperties["d"]["AP_DMS_Created"];
                                    }
                                    if ((string)jobjForProperties["d"]["AP_DMS_CreatedDate"] == null)
                                    {

                                    }
                                    else
                                    {
                                        objSharepointParam.AP_DMS_CreatedDate = (string)jobjForProperties["d"]["AP_DMS_CreatedDate"];
                                    }
                                    if ((string)jobjForProperties["d"]["AP_DMS_Modified_Date"] == null)
                                    {

                                    }
                                    else
                                    {
                                        objSharepointParam.AP_DMS_Modified_Date = (string)jobjForProperties["d"]["AP_DMS_Modified_Date"];
                                    }
                                    if ((string)jobjForProperties["d"]["AP_DMS_LASTMODIFIER"] == null)
                                    {

                                    }
                                    else
                                    {
                                        objSharepointParam.AP_DMS_LASTMODIFIER = (string)jobjForProperties["d"]["AP_DMS_LASTMODIFIER"];
                                    }
                                    if ((string)jobjForProperties["d"]["Reference_x0020_No"] == null)
                                    {

                                    }
                                    else
                                    {
                                        objSharepointParam.Reference_No = (string)jobjForProperties["d"]["Reference_x0020_No"];
                                    }
                                    if ((string)jobjForProperties["d"]["TDS_x0020_Amount"] == null)
                                    {

                                    }
                                    else
                                    {
                                        objSharepointParam.TDS_Amount = (string)jobjForProperties["d"]["TDS_x0020_Amount"];
                                    }
                                    if ((string)jobjForProperties["d"]["State"] == null)
                                    {

                                    }
                                    else
                                    {
                                        objSharepointParam.State = (string)jobjForProperties["d"]["State"];
                                    }
                                    if ((string)jobjForProperties["d"]["Class_x0020_ID"] == null)
                                    {

                                    }
                                    else
                                    {
                                        objSharepointParam.ClassID = (string)jobjForProperties["d"]["Class_x0020_ID"];
                                    }
                                    if ((string)jobjForProperties["d"]["CREATED_BY"] == null)
                                    {

                                    }
                                    else
                                    {
                                        objSharepointParam.CREATED_BY = (string)jobjForProperties["d"]["CREATED_BY"];
                                    }
                                    if ((string)jobjForProperties["d"]["CREATED_DATE"] == null)
                                    {

                                    }
                                    else
                                    {
                                        objSharepointParam.CREATED_DATE = (string)jobjForProperties["d"]["CREATED_DATE"];
                                    }
                                    if ((string)jobjForProperties["d"]["ENTITY_CODE"] == null)
                                    {

                                    }
                                    else
                                    {
                                        objSharepointParam.ENTITY_CODE = (string)jobjForProperties["d"]["ENTITY_CODE"];
                                    }
                                    if ((string)jobjForProperties["d"]["PROCESS_NUMBER"] == null)
                                    {

                                    }
                                    else
                                    {
                                        objSharepointParam.PROCESS_NUMBER = (string)jobjForProperties["d"]["PROCESS_NUMBER"];
                                    }
                                    if ((string)jobjForProperties["d"]["TYPE_ID"] == null)
                                    {

                                    }
                                    else
                                    {
                                        objSharepointParam.TYPE_ID = (string)jobjForProperties["d"]["TYPE_ID"];
                                    }
                                   
                                    ServicePointManager.SecurityProtocol = SecurityProtocolType.Tls | SecurityProtocolType.Tls11 | SecurityProtocolType.Tls12;
                                    using (WebClient client = new WebClient())
                                    {
                                        client.Headers.Add("X-FORMS_BASED_AUTH_ACCEPTED", "f");
                                        client.Headers.Add("Authorization", "Bearer " + objSharepointParam.access_token);
                                        Uri endpointUri = new Uri(uri + "/$value");
                                        byte[] dataByte = client.DownloadData(endpointUri);
                                        objSharepointParam.RelativeAttchmntPath = ServerRelativeUrl + "/" + folderName;
                                        string ServerRelativeUrlDest = listName;
                                        objSharepointParam.RelativeAttchmntPathDestination = ServerRelativeUrlDest + "/" + folderName;
                                        objSharepointParam.filename = fileName;
                                        UpdateDestinationPath(objSharepointParam, entityName);

                                        CopyfilesSharepoint(objSharepointParam, dataByte, objSharepointParam.filename);
                                    }

                                }
                            }
                        }
                    }
                    else
                    {
                        Console.WriteLine("List Name " + ListName + " could not be processed");
                    }
                }
                Console.WriteLine("getByteDataFromList Ended");
                #endregion
            }

            catch (Exception ex)
            {

                Console.WriteLine("getByteDataFromList Error");
            }
        }

        public void getByteDataFromList(SharepointParam objParam, string ServerRelativeUrl, string URL)
        {
            try
            {
                ServicePointManager.SecurityProtocol = SecurityProtocolType.Tls | SecurityProtocolType.Tls11 | SecurityProtocolType.Tls12;
                HttpWebRequest endpointRequestForList = (HttpWebRequest)HttpWebRequest.Create(objParam.sharepointaaccessurl + "_api/Web/GetFolderByServerRelativePath(decodedurl='" + objParam.ServerRelativeUrl + ServerRelativeUrl.ToString() + "')/Folders");
                HttpWebRequest endpointRequestForListTitle = (HttpWebRequest)HttpWebRequest.Create(objParam.sharepointaaccessurl + "_api/Web/GetFolderByServerRelativePath(decodedurl='" + objParam.ServerRelativeUrl + ServerRelativeUrl.ToString() + "')/Properties");

                endpointRequestForList.Method = "GET";
                endpointRequestForList.Accept = "application/json;odata=verbose";
                endpointRequestForList.Headers.Add("Authorization", "Bearer " + objParam.access_token);
                HttpWebResponse endpointResponseForList = (HttpWebResponse)endpointRequestForList.GetResponse();
                Stream webStreamList = endpointResponseForList.GetResponseStream();
                StreamReader responseReaderList = new StreamReader(webStreamList);
                string responseFileList = responseReaderList.ReadToEnd();
                JObject jobjList = JObject.Parse(responseFileList);
                JArray jarr = (JArray)jobjList["d"]["results"];

                endpointRequestForListTitle.Method = "GET";
                endpointRequestForListTitle.Accept = "application/json;odata=verbose";
                endpointRequestForListTitle.Headers.Add("Authorization", "Bearer " + objParam.access_token);
                HttpWebResponse endpointResponseForListTitle = (HttpWebResponse)endpointRequestForListTitle.GetResponse();
                Stream webStreamForListTitle = endpointResponseForListTitle.GetResponseStream();
                StreamReader responseReaderForListTitle = new StreamReader(webStreamForListTitle);
                string responseForListTitle = responseReaderForListTitle.ReadToEnd();
                JObject jobjForListTitle = JObject.Parse(responseForListTitle);
                string listName = (string)jobjForListTitle["d"]["vti_x005f_listtitle"];

                foreach (JObject k in jarr)
                {
                    string uriList = (string)k["Files"]["__deferred"]["uri"];
                    string folderName = (string)k["Name"];
                    if(folderName == "JayeshTest" || folderName == "Taxation")
                    {
                        Console.WriteLine("FolderName " + folderName);
                        ServicePointManager.SecurityProtocol = SecurityProtocolType.Tls | SecurityProtocolType.Tls11 | SecurityProtocolType.Tls12;
                        HttpWebRequest endpointRequestForFoldersItems = (HttpWebRequest)HttpWebRequest.Create(uriList);
                        endpointRequestForFoldersItems.Method = "GET";
                        endpointRequestForFoldersItems.Accept = "application/json;odata=verbose";
                        endpointRequestForFoldersItems.Headers.Add("Authorization", "Bearer " + objParam.access_token);
                        HttpWebResponse endpointResponseForFolderItems = (HttpWebResponse)endpointRequestForFoldersItems.GetResponse();
                        Stream webStreamFolderItems = endpointResponseForFolderItems.GetResponseStream();
                        StreamReader responseReaderFolderItems = new StreamReader(webStreamFolderItems);
                        string responseFileFolderItems = responseReaderFolderItems.ReadToEnd();
                        JObject jobjFolderItems = JObject.Parse(responseFileFolderItems);
                        JArray jarrFolder = (JArray)jobjFolderItems["d"]["results"];
                        foreach (JObject l in jarrFolder)
                        {
                            string uri = (string)l["__metadata"]["id"];
                            string ListItemsMetaData = (string)l["ListItemAllFields"]["__deferred"]["uri"];
                            string fileName = (string)l["Name"];
                            Console.WriteLine("Filename " + fileName);
                            ServicePointManager.SecurityProtocol = SecurityProtocolType.Tls | SecurityProtocolType.Tls11 | SecurityProtocolType.Tls12;
                            HttpWebRequest endpointRequestForListItems2 = (HttpWebRequest)HttpWebRequest.Create(uri + "/ListItemAllFields");
                            endpointRequestForListItems2.Method = "GET";
                            endpointRequestForListItems2.Accept = "application/json;odata=verbose";
                            endpointRequestForListItems2.Headers.Add("Authorization", "Bearer " + objParam.access_token);
                            HttpWebResponse endpointResponseForListItems2 = (HttpWebResponse)endpointRequestForListItems2.GetResponse();
                            Stream webStreamListItems2 = endpointResponseForListItems2.GetResponseStream();
                            StreamReader responseReaderListItems2 = new StreamReader(webStreamListItems2);
                            string responseFileListItems2 = responseReaderListItems2.ReadToEnd();
                            JObject jobjListItems2 = JObject.Parse(responseFileListItems2);
                            string entityName = "";
                            if ((string)jobjListItems2["d"]["EntityName"] == null)
                            {

                            }
                            else
                            {
                                entityName = (string)jobjListItems2["d"]["EntityName"];
                                objParam.Entity = (string)jobjListItems2["d"]["EntityName"];
                            }
                            if ((string)jobjListItems2["d"]["EntityGSTNumber"] == null)
                            {

                            }
                            else
                            {
                                objParam.Entity_GST_Number = (string)jobjListItems2["d"]["EntityGSTNumber"];
                            }
                            if ((string)jobjListItems2["d"]["FinYear"] == null)
                            {

                            }
                            else
                            {
                                objParam.Fin_Year = (string)jobjListItems2["d"]["FinYear"];
                            }
                            if ((string)jobjListItems2["d"]["Function"] == null)
                            {

                            }
                            else
                            {
                                objParam.Function = (string)jobjListItems2["d"]["Function"];
                            }
                            if ((string)jobjListItems2["d"]["InvoiceAmount"] == null)
                            {

                            }
                            else
                            {
                                objParam.Invoice_Amount = (string)jobjListItems2["d"]["InvoiceAmount"];
                            }
                            if ((string)jobjListItems2["d"]["InvoiceNumber"] == null)
                            {

                            }
                            else
                            {
                                objParam.Invoice_Number = (string)jobjListItems2["d"]["InvoiceNumber"];
                            }
                            if ((string)jobjListItems2["d"]["InvoiceDate"] == null)
                            {

                            }
                            else
                            {
                                objParam.Invoice_Date = (string)jobjListItems2["d"]["InvoiceDate"];
                            }
                            if ((string)jobjListItems2["d"]["Month"] == null)
                            {

                            }
                            else
                            {
                                objParam.Month = (string)jobjListItems2["d"]["Month"];
                            }
                            if ((string)jobjListItems2["d"]["PANNumber"] == null)
                            {

                            }
                            else
                            {
                                objParam.PAN_Number = (string)jobjListItems2["d"]["PANNumber"];
                            }
                            if ((string)jobjListItems2["d"]["PartyGSTNumber"] == null)
                            {

                            }
                            else
                            {
                                objParam.Party_GST_Number = (string)jobjListItems2["d"]["PartyGSTNumber"];
                            }
                            if ((string)jobjListItems2["d"]["ReferenceNumber_Name"] == null)
                            {

                            }
                            else
                            {
                                objParam.Reference_Number_Name = (string)jobjListItems2["d"]["ReferenceNumber_Name"];
                            }
                            if ((string)jobjListItems2["d"]["SubFunction"] == null)
                            {

                            }
                            else
                            {
                                objParam.Sub_Function = (string)jobjListItems2["d"]["SubFunction"];
                            }
                            if ((string)jobjListItems2["d"]["VendorName"] == null)
                            {

                            }
                            else
                            {
                                objParam.Vendor_Name = (string)jobjListItems2["d"]["VendorName"];
                            }
                            if ((string)jobjListItems2["d"]["AP_DMS_Created"] == null)
                            {

                            }
                            else
                            {
                                objParam.AP_DMS_Created = (string)jobjListItems2["d"]["AP_DMS_Created"];
                            }
                            if ((string)jobjListItems2["d"]["AP_DMS_CreatedDate"] == null)
                            {

                            }
                            else
                            {
                                objParam.AP_DMS_CreatedDate = (string)jobjListItems2["d"]["AP_DMS_CreatedDate"];
                            }
                            if ((string)jobjListItems2["d"]["AP_DMS_Modified_Date"] == null)
                            {

                            }
                            else
                            {
                                objParam.AP_DMS_Modified_Date = (string)jobjListItems2["d"]["AP_DMS_Modified_Date"];
                            }
                            if ((string)jobjListItems2["d"]["AP_DMS_LASTMODIFIER"] == null)
                            {

                            }
                            else
                            {
                                objParam.AP_DMS_LASTMODIFIER = (string)jobjListItems2["d"]["AP_DMS_LASTMODIFIER"];
                            }
                            if ((string)jobjListItems2["d"]["ReferenceNo"] == null)
                            {

                            }
                            else
                            {
                                objParam.Reference_No = (string)jobjListItems2["d"]["ReferenceNo"];
                            }
                            if ((string)jobjListItems2["d"]["TDSAmount"] == null)
                            {

                            }
                            else
                            {
                                objParam.TDS_Amount = (string)jobjListItems2["d"]["TDSAmount"];
                            }
                            if ((string)jobjListItems2["d"]["State"] == null)
                            {

                            }
                            else
                            {
                                objParam.State = (string)jobjListItems2["d"]["State"];
                            }
                            if ((string)jobjListItems2["d"]["ClassID"] == null)
                            {

                            }
                            else
                            {
                                objParam.ClassID = (string)jobjListItems2["d"]["ClassID"];
                            }
                            if ((string)jobjListItems2["d"]["CREATED_BY"] == null)
                            {

                            }
                            else
                            {
                                objParam.CREATED_BY = (string)jobjListItems2["d"]["CREATED_BY"];
                            }
                            if ((string)jobjListItems2["d"]["CREATED_DATE"] == null)
                            {

                            }
                            else
                            {
                                objParam.CREATED_DATE = (string)jobjListItems2["d"]["CREATED_DATE"];
                            }
                            if ((string)jobjListItems2["d"]["ENTITY_CODE"] == null)
                            {

                            }
                            else
                            {
                                objParam.ENTITY_CODE = (string)jobjListItems2["d"]["ENTITY_CODE"];
                            }
                            if ((string)jobjListItems2["d"]["PROCESS_NUMBER"] == null)
                            {

                            }
                            else
                            {
                                objParam.PROCESS_NUMBER = (string)jobjListItems2["d"]["PROCESS_NUMBER"];
                            }
                            if ((string)jobjListItems2["d"]["TYPE_ID"] == null)
                            {

                            }
                            else
                            {
                                objParam.TYPE_ID = (string)jobjListItems2["d"]["TYPE_ID"];
                            }
                            ServicePointManager.SecurityProtocol = SecurityProtocolType.Tls | SecurityProtocolType.Tls11 | SecurityProtocolType.Tls12;
                            using (WebClient client = new WebClient())
                            {
                                client.Headers.Add("X-FORMS_BASED_AUTH_ACCEPTED", "f");
                                client.Headers.Add("Authorization", "Bearer " + objParam.access_token);
                                Uri endpointUri = new Uri(uri + "/$value");
                                byte[] data = client.DownloadData(endpointUri);
                                objParam.RelativeAttchmntPath = ServerRelativeUrl + "/" + folderName;
                                string ServerRelativeUrlDest = fileName;
                                objParam.RelativeAttchmntPathDestination = ServerRelativeUrlDest + "/" + folderName;
                                objParam.filename = fileName;
                                UpdateDestinationPath(objParam, entityName);

                                CopyfilesSharepoint(objParam, data, objParam.filename);
                            }

                              
                        }
                    }
                }
            }
            catch(Exception ex)
            {

            }
        }

        public void UpdateDestinationPath (SharepointParam param,string entity)
        {
            try
            {
                if(entity.Contains("{{company_name1}}") || entity.Contains("{{company_name2}}"))
                {
                    param.client_id_destination = ConfigurationManager.AppSettings["CLIENT_ID_EBL"].ToString();
                    param.client_secret_destination = ConfigurationManager.AppSettings["CLIENT_SECRET_EBL"].ToString();
                    param.sitedomain_destination = ConfigurationManager.AppSettings["SITEDOMAIN_EBL"].ToString();
                    param.ServerRelativeUrl_destination = ConfigurationManager.AppSettings["LEADLIBRARYNAME_EBL"].ToString();
                    param.resource_destination= ConfigurationManager.AppSettings["RESOURCE_EBL"].ToString();
                    param.destinationsharepointurl = ConfigurationManager.AppSettings["DestinationSharePointURLEBL"].ToString();
                    param.tenantid_destination = ConfigurationManager.AppSettings["TENANTID_EBL"].ToString();
                }
                else if(entity.Contains("{{company_name3}}"))
                {
                    param.client_id_destination = ConfigurationManager.AppSettings["CLIENT_ID_EFSL"].ToString();
                    param.client_secret_destination = ConfigurationManager.AppSettings["CLIENT_SECRET_EFSL"].ToString();
                    param.sitedomain_destination = ConfigurationManager.AppSettings["SITEDOMAIN_EFSL"].ToString();
                    param.ServerRelativeUrl_destination = ConfigurationManager.AppSettings["LEADLIBRARYNAME_EFSL"].ToString();
                    param.resource_destination = ConfigurationManager.AppSettings["RESOURCE_EFSL"].ToString();
                    param.destinationsharepointurl = ConfigurationManager.AppSettings["DestinationSharePointURLEFSL"].ToString();
                    param.tenantid_destination = ConfigurationManager.AppSettings["TENANTID_EFSL"].ToString();
                }
                else if (entity.Contains("{{company_name4}}") || entity.Contains("{{company_name5}}") || entity.Contains("{{company_name6}}"))
                {
                    param.client_id_destination = ConfigurationManager.AppSettings["CLIENT_ID_EAML"].ToString();
                    param.client_secret_destination = ConfigurationManager.AppSettings["CLIENT_SECRET_EAML"].ToString();
                    param.sitedomain_destination = ConfigurationManager.AppSettings["SITEDOMAIN_EAML"].ToString();
                    param.ServerRelativeUrl_destination = ConfigurationManager.AppSettings["LEADLIBRARYNAME_EAML"].ToString();
                    param.resource_destination = ConfigurationManager.AppSettings["RESOURCE_EAML"].ToString();
                    param.destinationsharepointurl = ConfigurationManager.AppSettings["DestinationSharePointURLEAML"].ToString();
                    param.tenantid_destination = ConfigurationManager.AppSettings["TENANTID_EAML"].ToString();
                }
                else
                {
                    param.client_id_destination = ConfigurationManager.AppSettings["CLIENT_ID_EBL"].ToString();
                    param.client_secret_destination = ConfigurationManager.AppSettings["CLIENT_SECRET_EBL"].ToString();
                    param.sitedomain_destination = ConfigurationManager.AppSettings["SITEDOMAIN_EBL"].ToString();
                    param.ServerRelativeUrl_destination = ConfigurationManager.AppSettings["LEADLIBRARYNAME_EBL"].ToString();
                    param.resource_destination = ConfigurationManager.AppSettings["RESOURCE_EBL"].ToString();
                    param.destinationsharepointurl = ConfigurationManager.AppSettings["DestinationSharePointURLEBL"].ToString();
                    param.tenantid_destination = ConfigurationManager.AppSettings["TENANTID_EBL"].ToString();
                }
            }
            catch(Exception ex)
            {
                Console.WriteLine("Error in UpdateDestinationPath: " + ex.Message + "with stack trace " + ex.StackTrace);
            }
        }

        public static void UpdateFileProperties(SharepointParam ObjParams)
        {
            try
            {
                SharepointParam spParam = new SharepointParam();
                spParam.ObjFileProperty = new FileProperty();
                ObjParams.ObjFileProperty = new FileProperty();
                ObjParams.ObjFileProperty.__metadata = new CreateFolderMetadata();
                ObjParams.ObjFileProperty.__metadata.type = "SP.ListItem";
                string FolderFullPath = ObjParams.RelativeAttchmntPathDestination;
                FolderFullPath = FolderFullPath.Replace("\\", "/").Replace("//", "/");
                var folderUrls = FolderFullPath.Split(new char[] { '/' }, StringSplitOptions.RemoveEmptyEntries);
                ObjParams.FileFolder = string.Join("/", folderUrls, 0, folderUrls.Length - 1);
                var subFolderUrl = string.Join("/", folderUrls, 1, folderUrls.Length - 1);
                ObjParams.SubFolder = subFolderUrl;
                string URL = "";
                
                if (ObjParams.ServerRelativeUrl_destination.Contains(FolderFullPath))
                {
                    URL = ObjParams.destinationsharepointurl + "_api/web/GetFolderByServerRelativeUrl('" + ObjParams.ServerRelativeUrl_destination + "/" + ObjParams.filename + "')/ListItemAllFields";
                }
                else
                {
                    URL = ObjParams.destinationsharepointurl + "_api/web/GetFolderByServerRelativeUrl('" + ObjParams.ServerRelativeUrl_destination + FolderFullPath + "/" + ObjParams.filename + "')/ListItemAllFields";
                }
                Uri endpointUri = new Uri(URL);
                ObjParams.ListItem_Metadata_Type = "SP.ListItem";
                ObjParams.ObjFileProperty.__metadata.type = ObjParams.ListItem_Metadata_Type;
                ObjParams.ObjFileProperty.AP_DMS_Created = ObjParams.AP_DMS_Created;
                ObjParams.ObjFileProperty.AP_DMS_CreatedDate = ObjParams.AP_DMS_CreatedDate;
                ObjParams.ObjFileProperty.AP_DMS_LASTMODIFIER = ObjParams.AP_DMS_LASTMODIFIER;
                ObjParams.ObjFileProperty.AP_DMS_Modified_Date = ObjParams.AP_DMS_Modified_Date;
                ObjParams.ObjFileProperty.ClassID = ObjParams.ClassID;
                ObjParams.ObjFileProperty.CREATED_BY = ObjParams.CREATED_BY;
                ObjParams.ObjFileProperty.CREATED_DATE = ObjParams.CREATED_DATE;
                ObjParams.ObjFileProperty.Entity = ObjParams.Entity;
                ObjParams.ObjFileProperty.ENTITY_CODE = ObjParams.ENTITY_CODE;
                ObjParams.ObjFileProperty.Entity_GST_Number = ObjParams.Entity_GST_Number;
                ObjParams.ObjFileProperty.Fin_Year = ObjParams.Fin_Year;
                ObjParams.ObjFileProperty.Function = ObjParams.Function;
                ObjParams.ObjFileProperty.Invoice_Amount = ObjParams.Invoice_Amount;
                ObjParams.ObjFileProperty.Invoice_Date = ObjParams.Invoice_Date;
                ObjParams.ObjFileProperty.Invoice_Number = ObjParams.Invoice_Number;
                ObjParams.ObjFileProperty.Month = ObjParams.Month;
                ObjParams.ObjFileProperty.PAN_Number = ObjParams.PAN_Number;
                ObjParams.ObjFileProperty.Party_GST_Number = ObjParams.Party_GST_Number;
                ObjParams.ObjFileProperty.PROCESS_NUMBER = ObjParams.PROCESS_NUMBER;
                ObjParams.ObjFileProperty.Reference_No = ObjParams.Reference_No;
                ObjParams.ObjFileProperty.Reference_Number_Name = ObjParams.Reference_Number_Name;
                ObjParams.ObjFileProperty.State = ObjParams.State;
                ObjParams.ObjFileProperty.Sub_Function = ObjParams.Sub_Function;
                ObjParams.ObjFileProperty.TDS_Amount = ObjParams.TDS_Amount;
                ObjParams.ObjFileProperty.TYPE_ID = ObjParams.TYPE_ID;
                ObjParams.ObjFileProperty.Vendor_Name = ObjParams.Vendor_Name;
                ObjParams.ObjFileProperty.ClassID = ObjParams.ClassID;
                
                var jsonString = Newtonsoft.Json.JsonConvert.SerializeObject(ObjParams.ObjFileProperty);
                HttpWebRequest httpWReq = (HttpWebRequest)WebRequest.Create(endpointUri);

                //ServicePointManager.SecurityProtocol = (SecurityProtocolType)3072;
                Encoding encoding = new UTF8Encoding();
                httpWReq.ProtocolVersion = HttpVersion.Version11;
                httpWReq.Method = "POST";
                httpWReq.Accept = "application/json;odata=verbose";//("Accept-Encoding", "gzip, deflate, sdch");
                httpWReq.ContentType = "application/json;odata=verbose";
                httpWReq.Headers.Add("Authorization", "Bearer " + ObjParams.access_token_destination);
                httpWReq.Headers.Add("X-HTTP-Method", "MERGE");
                httpWReq.Headers.Add("IF-MATCH", "*");
                byte[] data = Encoding.ASCII.GetBytes(jsonString);
                httpWReq.ContentLength = data.Length;
                Stream requestStream = httpWReq.GetRequestStream();
                requestStream.Write(data, 0, data.Length);
                requestStream.Close();
                HttpWebResponse response = (HttpWebResponse)httpWReq.GetResponse();
                var responseString = new StreamReader(response.GetResponseStream()).ReadToEnd();
            }
            catch (Exception ex)
            {
                Console.WriteLine("Error in UpdateFileProperties: " + ex.Message + "with stack trace " + ex.StackTrace);
            }
        }

        public void AddSharpointFileProperties(SharepointParam ObjParams)
        {
            try
            {
                SharepointParam spParam = new SharepointParam();
                spParam.ObjFileProperty = new FileProperty();
                ObjParams.ObjFileProperty = new FileProperty();
                ObjParams.ObjFileProperty.__metadata = new CreateFolderMetadata();
                ObjParams.ObjFileProperty.__metadata.type = "SP.ListItem";
                string FolderFullPath = ObjParams.RelativeAttchmntPathDestination;
                string[] FolderNames = FolderFullPath.Split("/");
                string ListName = FolderNames[0];
                FolderFullPath = FolderFullPath.Replace("\\", "/").Replace("//", "/");
                var folderUrls = FolderFullPath.Split(new char[] { '/' }, StringSplitOptions.RemoveEmptyEntries);
                ObjParams.FileFolder = string.Join("/", folderUrls, 0, folderUrls.Length - 1);
                var subFolderUrl = string.Join("/", folderUrls, 1, folderUrls.Length - 1);
                ObjParams.SubFolder = subFolderUrl;
                string URL = "";
                string[] titleDestination = ObjParams.ServerRelativeUrl_destination.Split("/");
                string tileDst = titleDestination[3].ToString();

                if (ObjParams.ServerRelativeUrl_destination.Contains(FolderFullPath))
                {
                    URL = ObjParams.destinationsharepointurl + "_api/web/GetFolderByServerRelativeUrl('" + ObjParams.ServerRelativeUrl_destination + "/" + ObjParams.filename + "')/ListItemAllFields";
                    URL = "https:/{{sharepoint_url}}/sites/{{sharepoint_name}}/_api/web/Lists/getbytitle('{{filename}}')/fields";
                    URL = ObjParams.destinationsharepointurl + "_api/web/Lists/getbytitle('"+ tileDst + "')/fields";
                }
                else
                {
                    URL = ObjParams.destinationsharepointurl + "_api/web/GetFolderByServerRelativeUrl('" + ObjParams.ServerRelativeUrl_destination + FolderFullPath + "/" + ObjParams.filename + "')/ListItemAllFields";
                    URL = "https://{{sharepoint_url}}/sites/{{sharepoint_name}}/_api/web/Lists/getbytitle('{{filename}}')/fields";
                    URL = ObjParams.destinationsharepointurl + "_api/web/Lists/getbytitle('" + tileDst + "')/fields";
                }
                Uri endpointUri = new Uri(URL);
                ObjParams.ListItem_Metadata_Type = "SP.ListItem";
                ObjParams.ObjFileProperty.__metadata.type = ObjParams.ListItem_Metadata_Type;
                var jsonString = "";               
                string[] ListColumns = { "AP_DMS_Created", "AP_DMS_LASTMODIFIER", "Entity", "ENTITY_CODE",
                                        "Entity_GST_Number", "Fin_Year", "Function", "Invoice_Amount", "Invoice_Number",
                                        "Month","PAN_Number","Party_GST_Number","PROCESS_NUMBER","Reference_No",
                                        "Reference_Number_Name","State","Sub_Function","TDS_Amount","TYPE_ID","Vendor_Name","ClassID",
                                        "AP_DMS_Modified_Date", "AP_DMS_CreatedDate", "Invoice_Date", "CREATED_DATE"
                };
                foreach(string col in ListColumns)
                {
                    ServicePointManager.SecurityProtocol = SecurityProtocolType.Tls | SecurityProtocolType.Tls11 | SecurityProtocolType.Tls12;
                    jsonString = "{ '__metadata': { 'type': 'SP.Field' }, 'FieldTypeKind': 2,'Hidden':'false', 'Title':'" + col + "'}";
                    Encoding encoding = new UTF8Encoding();
                    HttpWebRequest httpWReq = (HttpWebRequest)WebRequest.Create(endpointUri);
                    httpWReq.ProtocolVersion = HttpVersion.Version11;
                    httpWReq.Method = "POST";
                    httpWReq.Accept = "application/json;odata=verbose";//("Accept-Encoding", "gzip, deflate, sdch");
                    httpWReq.ContentType = "application/json;odata=verbose";
                    httpWReq.Headers.Add("Authorization", "Bearer " + ObjParams.access_token_destination);
                    httpWReq.Headers.Add("X-RequestDigest", "#__REQUESTDIGEST");
                    byte[] data = Encoding.ASCII.GetBytes(jsonString);
                    httpWReq.ContentLength = data.Length;
                    Stream requestStream = httpWReq.GetRequestStream();
                    requestStream.Write(data, 0, data.Length);
                    requestStream.Close();
                    HttpWebResponse response = (HttpWebResponse)httpWReq.GetResponse();
                    var responseString = new StreamReader(response.GetResponseStream()).ReadToEnd();
                }
                

              
            }
            catch(Exception ex)
            {
                Console.WriteLine("Error In AddSharpointFileProperties:" + ex.Message + "-" + ex.StackTrace, "");
            }
        }

        public string CopyfilesSharepoint(SharepointParam objSharepointParam, byte[] data, string strFilePath)
        {
            try
            {
                int ListCount = 0;
                string responseToken = SharepointAccessForDestination(objSharepointParam);
                SharepointParam objResponsedest = JsonConvert.DeserializeObject<SharepointParam>(responseToken);

                objSharepointParam.access_token_destination = objResponsedest.access_token;
                string FolderPath = objSharepointParam.RelativeAttchmntPathDestination;
                string siteURL = "";
                if (objSharepointParam.ServerRelativeUrl_destination.EndsWith("/"))
                    siteURL = objSharepointParam.ServerRelativeUrl_destination.Remove(objSharepointParam.ServerRelativeUrl_destination.Length - 1);
                else
                    siteURL = objSharepointParam.ServerRelativeUrl_destination;

                string url = "";
                url = objSharepointParam.destinationsharepointurl + "_api/Web/GetFolderByServerRelativePath(decodedurl='" + siteURL + "')/folders";
                ServicePointManager.SecurityProtocol = SecurityProtocolType.Tls | SecurityProtocolType.Tls11 | SecurityProtocolType.Tls12;
                HttpWebRequest endpointRequestForFolderCheck = (HttpWebRequest)HttpWebRequest.Create(url);
                endpointRequestForFolderCheck.Method = "GET";
                endpointRequestForFolderCheck.Accept = "application/json;odata=verbose";
                endpointRequestForFolderCheck.Headers.Add("Authorization", "Bearer " + objSharepointParam.access_token_destination);
                HttpWebResponse endpointResponseForFolderCheck = (HttpWebResponse)endpointRequestForFolderCheck.GetResponse();
                Stream webStreamForFolderCheck = endpointResponseForFolderCheck.GetResponseStream();
                StreamReader responseReaderForFolderCheck = new StreamReader(webStreamForFolderCheck);
                string responseFileForFolderCheck = responseReaderForFolderCheck.ReadToEnd();
                JObject jobjForFolderCheck = JObject.Parse(responseFileForFolderCheck);
                JArray jarrFolder = (JArray)jobjForFolderCheck["d"]["results"];
                ListCount = jarrFolder.Count();
                if(ListCount == 1)
                {
                    AddSharpointFileProperties(objSharepointParam);
                }

                APICallSharepoint objApiCall = new APICallSharepoint();
                List<APIParameters> lstParameter = new List<APIParameters>();
                APIParameters objApiParameter = new APIParameters();
                objApiParameter.PropertyName = "Authorization";
                objApiParameter.PropertyValue = "Bearer " + objSharepointParam.access_token_destination;
                lstParameter.Add(objApiParameter);

                objApiParameter = new APIParameters();
                objApiParameter.PropertyName = "IF-MATCH";
                objApiParameter.PropertyValue = "*";
                lstParameter.Add(objApiParameter);

                string folderdata = SharepointFolderExistsForDestination(objSharepointParam);
                if (folderdata == "")
                {
                    folderdata = SharepointCreateFolderForFilesOnly(objSharepointParam);
                    FolderPath = "";
                }
                else
                {
                    FolderPath = FolderPath.Replace("\\", "/").Replace("//", "/");
                }
                string Url = objSharepointParam.destinationsharepointurl + "_api/web/GetFolderByServerRelativeUrl('" + objSharepointParam.ServerRelativeUrl_destination + FolderPath + "')/Files/add(url='" + objSharepointParam.filename + "', overwrite=true)";
                string response = objApiCall.httpwebrequest(Url, lstParameter, "", "application/json;odata=verbose", "Post", data);

                
                UpdateFileProperties(objSharepointParam);
                objSharepointParam.ServerRelativeUrl_destination = ConfigurationSettings.AppSettings["LEADLIBRARYNAME_DESTINATION"].ToString();
                return response;
            }
            catch (Exception ex)
            {
                throw ex;
            }
        }

    }
}

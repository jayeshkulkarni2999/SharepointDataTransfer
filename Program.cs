using Newtonsoft.Json;
using System;
using static SharepointTransfer.Sharepoint;

namespace SharepointTransfer
{
    class Program
    {
        static void Main(string[] args)
        {
            Program pg = new Program();

            try
            {
                Console.WriteLine("Sharepoint Transfer Utility Started");
                pg.ReadExcelFromSource();
                Console.WriteLine("Sharepoint Transfer Utility Ended");
            }
            catch (Exception ex)
            {
                //ExceptionLogging.SendErrorToText(ex);
                //MessageBox.Show(ex.Message.ToString());
            }
        }


        public void ReadExcelFromSource()
        {
            try
            {
                Console.WriteLine("ReadExcelFromSource Started");
                Sharepoint spo = new Sharepoint();
                SharepointParam spParam = new SharepointParam();
                spParam = spo.CreateSharepointObject();
                Console.WriteLine("Sharepoint Access for source start");
                string tokenresp = spo.SharepointAccess(spParam);
                Console.WriteLine("Access Token Received");
                SharepointParam objResponsetoken = JsonConvert.DeserializeObject<SharepointParam>(tokenresp);
                if (objResponsetoken != null)
                {
                    spParam.access_token = objResponsetoken.access_token;
                }

                Console.WriteLine("getSharepointFilesOnly Started");
                spo.getSharepointFilesOnly(spParam);//get files
                Console.WriteLine("getSharepointFilesOnly Ended");
            }
            catch(Exception ex)
            {

            }
        }


    }
}

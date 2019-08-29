using System;
using System.Collections.Generic;
using System.IO;
using System.Text;
using Windows.ApplicationModel.DataTransfer;
using Windows.Storage;

namespace LangWars
{
    static class FileShare
    {
        public static void Register()
        {
            var dtm = DataTransferManager.GetForCurrentView();
            dtm.DataRequested += dtm_DataRequested;
        }

        public static void Share()
        {
            DataTransferManager.ShowShareUI();
        }

        public static async void dtm_DataRequested(DataTransferManager sender, DataRequestedEventArgs args)
        {
            try
            {
                args.Request.Data.Properties.Title = "Language Wars";
                args.Request.Data.Properties.Description = "Send the file with statistics to another app.";
                args.Request.Data.Properties.FileTypes.Add(".xlsx");
                var file = await ReportGenerator.TempXlsPath.GetFileAsync(ReportGenerator.TempXlsName);
                args.Request.Data.SetStorageItems(new IStorageFile[] { file });
            }
            catch (FileNotFoundException)
            {
                args.Request.FailWithDisplayText("There is no generated file to share. Make sure to press the 'Fight' button first.");
            }
            catch (Exception ex)
            {
                args.Request.FailWithDisplayText("There was an error: " + ex.Message);
            }
        }

    }
}

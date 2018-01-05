using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace PDB_Excel_Data_Extractor
{
    class SpreadSheetToXlsx
    {
        public void export()
        {
            string jsScript = @"function exportAsExcel(spreadsheetId) {
            var file = Drive.Files.get(spreadsheetId);
            var url = file.exportLinks['application/vnd.openxmlformats-officedocument.spreadsheetml.sheet'];
            var token = ScriptApp.getOAuthToken();
            var response = UrlFetchApp.fetch(url, {
                headers:
                {
                    'Authorization': 'Bearer ' + token
                }
            });
            return response.getBlob();
        }";
        }
    }
}

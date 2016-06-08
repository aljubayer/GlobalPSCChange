using System;
using System.Collections;
using System.Collections.Generic;
using System.Data;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using ManiacProject.Libs;

namespace GlobalPSC
{
    internal class Neighbor_Delete
    {
        public string RNCName { set; get; }
        public string RNCID { set; get; }
        public string CellID { set; get; }
        public string RNCIDofNeighboringCell { set; get; }
        public string NeighboringCellID { set; get; }


    }

    internal class PSCInfo
    {
        public string RNCName { set; get; }
        public string CELLID { set; get; }
        public string CELLNAME { set; get; }
        public string NEWPSC { set; get; }
    }

    internal class External3G_Update
    {
        public string RNCName { set; get; }
        public string NeighboringRNCID { set; get; }
        public string CellIDofNeighboringRNC { set; get; }
        public string PSCRAMBCODE { set; get; }

    }

    public class HuaweiPSCByCME
    {
        private string inputFile = string.Empty;
        private string templateFile = string.Empty;
        private List<PSCInfo> pscInfo = new List<PSCInfo>();
        private List<Neighbor_Delete> neighborDelete = new List<Neighbor_Delete>();
        private List<External3G_Update> external3GUpdate = new List<External3G_Update>();
        Dictionary<string,List<Dictionary<string, string>>> notFound = new Dictionary<string, List<Dictionary<string, string>>>();
            
        Dictionary<string,List<Dictionary<string,string>>> templateData = new Dictionary<string, List<Dictionary<string, string>>>(); 

        public string ProcessTemplate(string inputFile, string templateFile)
        {
            this.inputFile = inputFile;
            this.templateFile = templateFile;
            ReadInputFile();
            ReadTemplateFile();
            UpdateTemplateData();
            string curr = Directory.GetCurrentDirectory();
            IOFileOperation.CreateExelFile(templateData, curr);

            if (File.Exists(curr + @"\DataNotFound.xlsx"))
            {
                File.Delete(curr + @"\DataNotFound.xlsx");
            }
            if (notFound.Count != 0)
            {
                IOFileOperation.CreateExelFile(notFound, curr, "DataNotFound.xlsx");
            }
          
            return "Finished";
        }

        private void UpdateTemplateData()
        {
            List<Dictionary<string, string>> notFoundPscInfo = new List<Dictionary<string, string>>();
            foreach (PSCInfo info in pscInfo)
            {
                if (templateData["CELL"].Exists(i => i["BSCName"].Trim() == info.RNCName.Trim() && i["CELLID"].Trim() == info.CELLID.Trim()))
                {
                    int index =
                        templateData["CELL"].FindIndex(i => i["BSCName"].Trim() == info.RNCName.Trim() && i["CELLID"].Trim() == info.CELLID.Trim());

                    templateData["CELL"][index]["PSCRAMBCODE"] = info.NEWPSC;

                }
                else
                {
                    Dictionary<string,string> aDictionary = new Dictionary<string, string>();
                    aDictionary.Add("RNCName", info.RNCName);
                    aDictionary.Add("CELLID", info.CELLID);
                    aDictionary.Add("CELLNAME", info.CELLNAME);
                    aDictionary.Add("NEWPSC", info.NEWPSC);
                    notFoundPscInfo.Add(aDictionary);
                }
            }

            if (notFoundPscInfo.Count != 0)
            {
                notFound.Add("PSCInfo", notFoundPscInfo);
            }

            List<Dictionary<string, string>> notFoundneighborDelete = new List<Dictionary<string, string>>();
            foreach (Neighbor_Delete delete in neighborDelete)
            {
                if (templateData["INTRAFREQNCELL"].Exists(i => i["BSCName"].Trim() == delete.RNCName.Trim()
                                                               && i["CELLID"].Trim() == delete.CellID.Trim()
                                                               && i["NCELLRNCID"].Trim() == delete.RNCIDofNeighboringCell.Trim()
                                                               && i["NCELLID"].Trim() == delete.NeighboringCellID.Trim()))
                {
                    int index =
                        templateData["INTRAFREQNCELL"].FindIndex(i => i["BSCName"].Trim() == delete.RNCName.Trim()
                                                                      && i["CELLID"].Trim() == delete.CellID.Trim()
                                                                      && i["NCELLRNCID"].Trim() == delete.RNCIDofNeighboringCell.Trim()
                                                                      && i["NCELLID"].Trim() == delete.NeighboringCellID.Trim());
                    Dictionary<string, string> aDictionary = new Dictionary<string, string>();
                    foreach (KeyValuePair<string, string> keyValuePair in templateData["INTRAFREQNCELL"][index])
                    {
                        aDictionary.Add(keyValuePair.Key, "");
                    }

                    templateData["INTRAFREQNCELL"][index] = aDictionary;
                }
                else
                {
                    Dictionary<string,string> aDictionary = new Dictionary<string, string>();
                    aDictionary.Add("RNCName", delete.RNCName);
                    aDictionary.Add("RNCID", delete.RNCID);
                    aDictionary.Add("CellID", delete.CellID);
                    aDictionary.Add("RNCIDofNeighboringCell", delete.RNCIDofNeighboringCell);
                    aDictionary.Add("NeighboringCellID", delete.NeighboringCellID);
                    notFoundneighborDelete.Add(aDictionary);
                 
                }
            }

            if (notFoundneighborDelete.Count != 0)
            {
                notFound.Add("Neighbor_Delete(3G-3G)", notFoundneighborDelete);
            }


            List<Dictionary<string,string>> notFoundExternalUpdate = new List<Dictionary<string, string>>();
            foreach (External3G_Update gUpdate in external3GUpdate)
            {
                if (templateData["NRNCCELL"].Exists(i => i["BSCName"].Trim() == gUpdate.RNCName.Trim()
                    && i["NRNCID"].Trim() == gUpdate.NeighboringRNCID.Trim()
                    && i["CELLID"].Trim() == gUpdate.CellIDofNeighboringRNC.Trim()))
                {
                    int index =
                        templateData["NRNCCELL"].FindIndex(i => i["BSCName"].Trim() == gUpdate.RNCName.Trim()
                                                                && i["NRNCID"].Trim() == gUpdate.NeighboringRNCID.Trim()
                                                                && i["CELLID"].Trim() == gUpdate.CellIDofNeighboringRNC.Trim());

                    templateData["NRNCCELL"][index]["PSCRAMBCODE"] = gUpdate.PSCRAMBCODE;

                }
                else
                {
                    Dictionary<string,string> aDictionary = new Dictionary<string, string>();
                    aDictionary.Add("RNC:", gUpdate.RNCName);
                    aDictionary.Add("NeighboringRNCID:", gUpdate.NeighboringRNCID);
                    aDictionary.Add("CellIDofNeighboringRNC:", gUpdate.CellIDofNeighboringRNC);
                    aDictionary.Add("PSCRAMBCODE:", gUpdate.PSCRAMBCODE);
                    notFoundExternalUpdate.Add(aDictionary);
                }
            }

            if (notFoundExternalUpdate.Count != 0)
            {
                notFound.Add("External3G_Update",notFoundExternalUpdate);
            }
        }

        private void ReadTemplateFile()
        {
            List<Dictionary<string, string>> CELL = new List<Dictionary<string, string>>();
            DataSet aSet = IOFileOperation.ReadExcelMacroFile(templateFile, "CELL");

            List<string> cols = new List<string>();
            foreach (DataColumn dataColumn in aSet.Tables[0].Columns)
            {
                cols.Add(dataColumn.ColumnName);
            }
            foreach (DataRow dataRow in aSet.Tables[0].Rows)
            {
                Dictionary<string, string> aDictionary = new Dictionary<string, string>();
                foreach (string col in cols)
                {
                    aDictionary.Add(col, dataRow[col].ToString());
                }

                CELL.Add(aDictionary);
            }
            templateData.Add("CELL", CELL);



            List<Dictionary<string, string>> NRNCCELL = new List<Dictionary<string, string>>();
            aSet = IOFileOperation.ReadExcelMacroFile(templateFile, "NRNCCELL");

            cols = new List<string>();
            foreach (DataColumn dataColumn in aSet.Tables[0].Columns)
            {
                cols.Add(dataColumn.ColumnName);
            }
            foreach (DataRow dataRow in aSet.Tables[0].Rows)
            {
                Dictionary<string, string> aDictionary = new Dictionary<string, string>();
                foreach (string col in cols)
                {
                    aDictionary.Add(col, dataRow[col].ToString());
                }

                NRNCCELL.Add(aDictionary);
            }
            templateData.Add("NRNCCELL", NRNCCELL);



            List<Dictionary<string, string>> INTRAFREQNCELL = new List<Dictionary<string, string>>();
            aSet = IOFileOperation.ReadExcelMacroFile(templateFile, "INTRAFREQNCELL");

             cols = new List<string>();
            foreach (DataColumn dataColumn in aSet.Tables[0].Columns)
            {
                cols.Add(dataColumn.ColumnName);
            }
            foreach (DataRow dataRow in aSet.Tables[0].Rows)
            {
                Dictionary<string,string> aDictionary = new Dictionary<string, string>();
                foreach (string col in cols)
                {
                    aDictionary.Add(col,dataRow[col].ToString());
                }

                INTRAFREQNCELL.Add(aDictionary);
            }
            templateData.Add("INTRAFREQNCELL", INTRAFREQNCELL);



           

        }


        private void ReadInputFile()
        {
            DataSet aSet = IOFileOperation.ReadExcelFile(inputFile, "PSCInfo");

            foreach (DataRow dataRow in aSet.Tables[0].Rows)
            {
                PSCInfo aPscInfo = new PSCInfo();
                aPscInfo.RNCName = dataRow["RNCName"].ToString().Trim();
                aPscInfo.CELLID = dataRow["CELLID"].ToString().Trim();
                aPscInfo.CELLNAME = dataRow["CELLNAME"].ToString().Trim();
                aPscInfo.NEWPSC = dataRow["NEW PSC"].ToString().Trim();
                if (aPscInfo.RNCName.Trim().Length != 0)
                {
                    pscInfo.Add(aPscInfo);
                }

            }


            aSet = IOFileOperation.ReadExcelFile(inputFile, "Neighbor_Delete(3G-3G)");
            foreach (DataRow dataRow in aSet.Tables[0].Rows)
            {
                Neighbor_Delete aNeighborDelete = new Neighbor_Delete();
                aNeighborDelete.RNCName = dataRow["RNCName"].ToString().Trim();
                aNeighborDelete.RNCID = dataRow["RNC ID"].ToString().Trim();
                aNeighborDelete.CellID = dataRow["Cell ID"].ToString().Trim();
                aNeighborDelete.RNCIDofNeighboringCell = dataRow["RNC ID of a neighboring cell"].ToString().Trim();
                aNeighborDelete.NeighboringCellID = dataRow["Neighboring Cell ID"].ToString().Trim();

                if (aNeighborDelete.RNCName.Trim().Length != 0)
                {
                    neighborDelete.Add(aNeighborDelete);
                }
            }
            aSet = IOFileOperation.ReadExcelFile(inputFile, "External3G_Update");

            foreach (DataRow dataRow in aSet.Tables[0].Rows)
            {
                External3G_Update aExternal3GUpdate = new External3G_Update();
                aExternal3GUpdate.RNCName = dataRow["RNCName"].ToString().Trim();
                aExternal3GUpdate.NeighboringRNCID = dataRow["Neighboring RNC ID"].ToString().Trim();
                aExternal3GUpdate.CellIDofNeighboringRNC = dataRow["Cell ID of Neighboring RNC"].ToString().Trim();
                aExternal3GUpdate.PSCRAMBCODE = dataRow["PSCRAMBCODE"].ToString().Trim();

                if (aExternal3GUpdate.RNCName.Trim().Length != 0)
                {
                    external3GUpdate.Add(aExternal3GUpdate);
                }
            }


        }
    }
}

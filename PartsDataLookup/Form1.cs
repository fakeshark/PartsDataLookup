using System;
using System.Collections.Generic;
using System.Data;
using System.Linq;
using System.Text;
using System.Windows.Forms;
using System.IO;
using ExcelDataReader;
using OfficeOpenXml;

namespace PartsDataLookup
{
    public partial class Form1 : Form
    {
        #region Globals
        string[] IndcArray;
        string[] CageCodeArray;
        string[] PartNumsArray;
        string[] NamesArray;
        string[] FscArray;
        string[] NiinArray;
        string[] NhaArray;
        string[] QpaArray;
        string[] ProvNomArray;
        string[] NotesArray;
        int ouptutIndex = 0;
        int exactMatches = 0;
        int pnMatches = 0;
        int multiMatches = 0;
        int nonMatches = 0;
        DataTable table = new DataTable();
        List<string[]> MatchList = new List<string[]>() { };
        List<string[]> PartsFrom036 = new List<string[]>() { };
        #endregion

        public Form1()
        {
            InitializeComponent();
        }

        private void LoadPartsSearchList_Click(object sender, EventArgs e)
        {
            LoadSearchList();
        }

        public void LoadSearchList()
        {
            OpenFileDialog loadNSNList = new OpenFileDialog()
            {
                Filter = "Excel Spread Sheets|*.xlsx|Text Files|*.txt|All Files|*.*",
                Title = "Select an NSN List File",
                Multiselect = false
            };

            if (loadNSNList.ShowDialog() == System.Windows.Forms.DialogResult.OK)
            {
                string fileExtension = Path.GetExtension(loadNSNList.FileName);

                //Todo: add handling for a text file list
                switch (fileExtension)
                {
                    case ".xlsx":
                        ExcelListOpen(loadNSNList);
                        break;

                    //case ".txt":
                    //    MessageBox.Show("Text File");
                    //    break;

                    default:
                        MessageBox.Show("Sorry, handling for this file type (" + fileExtension + ") has not yet been developed.");
                        break;
                }
            }
            else
            {
                MessageBox.Show("File Load Cancelled");
            }
        }

        public void ExcelListOpen(OpenFileDialog loadNSNList)
        {
            string filePath = loadNSNList.FileName;

            IExcelDataReader excelReader;
            FileStream stream = File.Open(filePath, FileMode.Open, FileAccess.Read);

            excelReader = ExcelReaderFactory.CreateOpenXmlReader(stream);
            DataSet result = excelReader.AsDataSet();

            result = excelReader.AsDataSet(new ExcelDataSetConfiguration()
            {
                UseColumnDataType = true,

                // Gets or sets a callback to obtain configuration options for a DataTable. 
                ConfigureDataTable = (tableReader) => new ExcelDataTableConfiguration()
                {
                    // Gets or sets a value indicating the prefix of generated column names.
                    EmptyColumnNamePrefix = "Column",

                    // Gets or sets a value indicating whether to use a row from the data as column names.
                    UseHeaderRow = true
                }
            });

            dgvExcelList.DataSource = result.Tables[0];
            dgvExcelList.Update();

            string columns = result.Tables[0].Columns.Count.ToString();
            string rows = result.Tables[0].Rows.Count.ToString();

            List<string> IndcList = new List<string>() { };
            List<string> cageCodeList = new List<string>() { };
            List<string> partNumsList = new List<string>() { };
            List<string> namesList = new List<string>() { };
            List<string> fscList = new List<string>() { };
            List<string> niinList = new List<string>() { };
            List<string> nhaList = new List<string>() { };
            List<string> qpaList = new List<string>() { };
            List<string> provNomList = new List<string>() { };
            List<string> notesList = new List<string>() { };

            int rowCounter = 0;

            foreach (DataGridViewRow listRow in dgvExcelList.Rows)
            {
                rowCounter++;
                // only add part numbers that actually contain values
                if (listRow.Cells[4].Value != null)
                {
                    if (listRow.Cells[4].Value.ToString().Trim() != "")
                    {
                        IndcList.Add(listRow.Cells[2].Value.ToString().Trim());
                        cageCodeList.Add(listRow.Cells[3].Value.ToString().Trim());
                        partNumsList.Add(listRow.Cells[4].Value.ToString().Trim());
                        namesList.Add(listRow.Cells[5].Value.ToString().Trim());
                        fscList.Add(listRow.Cells[6].Value.ToString().Trim());
                        niinList.Add(listRow.Cells[7].Value.ToString().Trim());
                        nhaList.Add(listRow.Cells[8].Value.ToString().Trim());
                        qpaList.Add(listRow.Cells[9].Value.ToString().Trim());
                        provNomList.Add(listRow.Cells[10].Value.ToString().Trim());
                        notesList.Add(listRow.Cells[11].Value.ToString().Trim() + "; " + listRow.Cells[12].Value.ToString().Trim());
                    }
                }
            }

            IndcArray = IndcList.ToArray();
            CageCodeArray = cageCodeList.ToArray();
            PartNumsArray = partNumsList.ToArray();
            NamesArray = namesList.ToArray();
            FscArray = fscList.ToArray();
            NiinArray = niinList.ToArray();
            NhaArray = nhaList.ToArray();
            QpaArray = qpaList.ToArray();
            ProvNomArray = provNomList.ToArray();
            NotesArray = notesList.ToArray();

            //PartNumsArray = partNumsList.Distinct().ToArray();

            lblPartsToMatch.Text = rowCounter + " rows loaded and " + PartNumsArray.Count() + " usable parts identified.";
            btnLoadFlisFoi.Enabled = true;
        }

        private void LoadFlisFoi_Click(object sender, EventArgs e)
        {
            if (PartNumsArray != null && PartNumsArray.Length != 0)
            {
                LoadFlisList(PartNumsArray);
            }
            btnMatchRecords.Enabled = true;
        }

        private void LoadFlisList(string[] partNumsArray)
        {
            OpenFileDialog loadFlisTextFile = new OpenFileDialog();

            loadFlisTextFile.Filter = "Text Files|*.txt";
            loadFlisTextFile.Title = "Select a FLIS Packaging Data File";
            IEnumerable<string> fileData;

            if (loadFlisTextFile.ShowDialog() == System.Windows.Forms.DialogResult.OK)
            {
                string flisPath = loadFlisTextFile.FileName;
                fileData = File.ReadLines(flisPath);

                List<string> TempMatchPart = new List<string>() { };
                bool partFound = false;

                foreach (string item in fileData)
                {
                    if (partFound == false)
                    {
                        // if '01' line found, add to tempMatchPart list and partFound becomes true
                        if (item.Substring(0, 2) == "01")
                        {
                            TempMatchPart.Add(item);
                            partFound = true;
                        }
                    }
                    else if (partFound == true)
                    {
                        // once partFound is true, keep adding lines until another '01' line is found
                        if (item.Substring(0, 2) != "01")
                        {
                            TempMatchPart.Add(item);
                        }
                        else
                        {
                            // another '01' line is discovered, so the last partFound is done
                            // check the last part to see if it contains any of our part numbers
                            // if it does, add it to the MatchPart list. if not, do nothing with it
                            // then clear out TempMatchPart and start again with this new '01' line
                            CheckForPartNumbers(TempMatchPart);
                            TempMatchPart.Clear();
                            TempMatchPart.Add(item);
                            partFound = true;
                        }
                    }
                }
                lblMatchingRecords.Text = MatchList.Count.ToString() + " records matched from " + string.Format("{0:n0}", fileData.Count()) + " lines of data.";
            }
        }

        private void CheckForPartNumbers(List<string> tempMatchPart)
        {
            bool PartNumberMatched = false;
            int matchCounter = 0;
            // look for '03' line and check if any of our part numbers are in there in the given tempMatchPart List
            foreach (string line in tempMatchPart)
            {
                if (line.Substring(0, 2) == "03")
                {
                    // if any of our part numbers are present on any of the '03' lines, increase the match counter by one
                    if (PartNumsArray.Contains(line.Substring(15, 31).Trim()))
                    {
                        PartNumberMatched = true;
                        matchCounter++;
                    }
                }
            }

            // if match counter is higher than zero, that indicates that at least one of our PN's have been found
            if (PartNumberMatched)
            {
                // make sure we havent already added this part array to the list
                if (!MatchList.Contains(tempMatchPart.ToArray()))
                {
                    // add the part array to the list or matching part arrays
                    MatchList.Add(tempMatchPart.ToArray());
                }                
            }
        }

        private void MatchRecords_Click(object sender, EventArgs e)
        {
            string[] pplHeaderRow = new string[] {"INDEX",  "PCCN", "PLISN", "INDC", "CAGE", "PN", "RNCC", "RNVC", "DAC", "PPSL", "EC", "NAME", "SL", "SLAC", "COG", "MCC", "FSC", "NIIN", "NSNSUFF", "UM", "UM PRICE", "UI", "UI PRICE", "CONV", "QUP", "SMR", "DMIL", "PLT", "HCI", "PSPC", "PMIC", "ADPEC", "NHA", "ORR", "QPA", "QPE", "MRRI", "MRRII", "MRR MOD", "TQR", "SAPLISN", "PRPLISN", "MAOT", "MAC", "NRTS", "UOC", "REFDES", "RDOC", "RDC", "SMCC", "PLCC", "SMIC", "AIC", "AIC QTY", "MRU", "RMSS", "RISS", "RTLL QTY", "RSR", "O-MTD", "F-MTD", "H-MTD", "SRA-MTD", "D-MTD", "CED-MTD", "CAD-MTD", "O-RCT", "F-RCT", "H-RCT", "SRA-RCT", "D-RCT", "CON-RCT", "O-RTD", "F-RTD", "H-RTD", "SRA-RTD", "D-RTD", "DOP1", "DOP2", "CTIC", "AMC", "AMSC", "IMC", "RIP", "CHANGE AUTHORITY1", "IC", "SN FROM", "SN TO", "TIC", "R/S PLISN", "QTY SHIPPED", "QTY PROCURED", "DCN UOC", "PRORATED ELIN", "PRORATED QTY", "LCN", "ALT LCN", "REMARKS", "TM CODE", "FIG NUM", "ITEM NUM", "TM CHG", "TM IND", "QTY FIG", "WUC/TM FGC", "BASIS OF ISSUE1", "BASIS OF ISSUE2", "CC", "INC", "LRU", "PROV NOM", "ALTCAGE1", "ALTPN1", "ALTRNCC1", "ALTRNVC1", "ALTDAC1", "ALTPPSL1", "ALTCAGE2", "ALTPN2", "ALTRNCC2", "ALTRNVC2", "ALTDAC2", "ALTPPSL2", "ALTCAGE3", "ALTPN3", "ALTRNCC3", "ALTRNVC3", "ALTDAC3", "ALTPPSL3", "ALTCAGE4", "ALTPN4", "ALTRNCC4", "ALTRNVC4", "ALTDAC4", "ALTPPSL4", "ALTCAGE5", "ALTPN5", "ALTRNCC5", "ALTRNVC5", "ALTDAC5", "ALTPPSL5", "ALTCAGE6", "ALTPN6", "ALTRNCC6", "ALTRNVC6", "ALTDAC6", "ALTPPSL6", "MATERIAL1", "MATERIAL2", "MATERIAL3", "MATERIAL4", "RBD", "SUPP NOMEN 1", "AEL-A", "AEL-B", "AEL-C", "AEL-D", "AEL-E", "AEL-F", "AEL-G", "AEL-H", "SUPP NOMEN 2", "AFC2", "AFC QTY2", "ANC2", "AOC2", "AOC QTY2", "LLTIL1", "PPL1", "SFPPL1", "CBIL1", "RIL1", "ISIL1", "PCL1", "TTEL1", "SCPL1", "DCN1", "ARF", "LLTIL2", "PPL2", "SFPPL2", "CBIL2", "RIL2", "ISIL2", "PCL2", "TTEL2", "SCPL2", "DCN2", "ACC CODE", "ALT NIIN REL", "ALT NIIN", "ALT NIIN REL2", "ALT NIIN2", "REFDES2", "RDOC2", "CHANGE AUTHORITY2", "IC2", "SN FROM2", "SN TO2", "TIC2", "R/S PLISN2", "QTY SHIPPED2", "QTY PROCURED2", "DCN UOC2", "PRORATED ELIN2", "PRORATED QTY2", "LCN2", "ALT LCN2", "LENGTH", "WIDTH", "HEIGHT", "WEIGHT", "temp1" };
            List<string[]> pplDataRow = new List<string[]>() { };
            
            for (int i = 0; i < PartNumsArray.Length; i++)
            {
                pplDataRow.Add(MakePplRow(CageCodeArray[i], PartNumsArray[i], NamesArray[i], i));
            }
                        
            for (int i = 0; i < pplHeaderRow.Length; i++)
            {
                // set the correct data type for each of these fields eventually
                table.Columns.Add(pplHeaderRow[i], typeof(string));
            }

            foreach (string[] row in pplDataRow)
            {
                table.Rows.Add(row);
            }

            dgvMatchList.DataSource = table;
            dgvMatchList.AutoResizeColumns();

            lblMatchResults.Text = exactMatches.ToString() + " exact matches identified.\n" + pnMatches.ToString() + " items matched part number only and require attention.\n" + multiMatches.ToString() + " items matching multiple records.\n" + nonMatches.ToString() + " items where no match could be identified.";
            btnExportToExcell.Enabled = true;
        }

        private string[] MakePplRow(string cage, string pn, string name, int pnIndex)
        {
            string[] pplColumns = new string[] { "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", " ", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "" };
            ouptutIndex++;
            List<string[]> matchingParts = new List<string[]>() { };
            string line01 = string.Empty;
            string line02 = string.Empty;
            string TEMP1 = string.Empty;

            // check every part array in the list of matching part arrays
            // first pass is to check for exact matches
            foreach (string[] partArray in MatchList)
            {
                for (int i = 0; i < partArray.Length; i++)
                {
                    if (partArray[i].Substring(0, 2) == "03")
                    {
                        if (partArray[i].Substring(15, 31).Trim().ToUpper() == pn.Trim().ToUpper() && partArray[i].Substring(10, 5) == cage)
                        {
                            if (!matchingParts.Contains(partArray))
                            {
                                matchingParts.Add(partArray);
                                TEMP1 = "Part, Cage: " + pn + ", " + cage + " found an exact match. User supplied name: (" + NamesArray[pnIndex].ToUpper() + ")    FLIS name: (";
                            }
                        }
                    }
                }
            }

            // if no exact matches are found search for pn only matches
            if (matchingParts.Count == 0)
            {
                foreach (string[] partArray in MatchList)
                {
                    for (int i = 0; i < partArray.Length; i++)
                    {
                        if (partArray[i].Substring(0, 2) == "03")
                        {
                            if ((partArray[i].Substring(15, 31).Trim().ToUpper() == pn.Trim().ToUpper() || StripPunctuation(partArray[i].Substring(15, 31).Trim().ToUpper()) == StripPunctuation(pn.Trim().ToUpper())))
                            {
                                if (!matchingParts.Contains(partArray))
                                {
                                    matchingParts.Add(partArray);
                                }
                            }
                        }
                    }
                }
            }

            List<string> unsorted03Lines = new List<string>() { };
            string best03Line = string.Empty;
            List<string> unsorted05Lines = new List<string>() { };
            string best05Line = string.Empty;

            // try to get multiple matches down to a single match by eliminating less than ideal options
            if (matchingParts.Count > 1)
            {
                matchingParts = PickBestPart(matchingParts, pn, cage);
            }

            if (matchingParts.Count == 1)
            {
                foreach (string[] arr in matchingParts)
                {
                    for (int i = 0; i < arr.Length; i++)
                    {
                        if (arr[i].Substring(0, 2) == "01")
                        {
                            line01 = arr[i];
                        }
                        else if (arr[i].Substring(0, 2) == "02")
                        {
                            line02 = arr[i];
                        }
                        else if (arr[i].Substring(0, 2) == "03")
                        {
                            unsorted03Lines.Add(arr[i]);
                        }
                        else if (arr[i].Substring(0, 2) == "05")
                        {
                            unsorted05Lines.Add(arr[i]);
                        }
                    }
                }
                if (TEMP1 == "")
                {
                    TEMP1 = "Part Number: " + pn + " found a possible match but needs user confirmation. Please compare user supplied name: (" + NamesArray[pnIndex].ToUpper() + ") with FLIS name: (";
                    pnMatches++;
                }
                best03Line = Sort03Lines(unsorted03Lines);
                best05Line = Sort05Lines(unsorted05Lines);                
            }
            else if (matchingParts.Count > 1)
            {
                TEMP1 = "Multiple (" + matchingParts.Count + ") data sections matched Part, Cage: " + pn + ", " + cage;
                multiMatches++;
            }
            else if (matchingParts.Count == 0)
            {
                TEMP1 = "No FLIS Data Found for Part, Cage: " + pn + ", " + cage;
                nonMatches++;
            }

            // CAGE & PN
            string CAGE = string.Empty;
            string PN = string.Empty;
            string PROVNOM = string.Empty;
            string REMARKS = string.Empty;
            if (best03Line.Length >= 16)
            {
                CAGE = best03Line.Substring(10, 5);
                PN = best03Line.Substring(15, 31).Trim();
                if (PN != pn)
                {
                    REMARKS = "PN \"" + PN + "\" has been updated, the original PN was: \"" + pn + "\"";
                }
            }

            // RNCC, RNVC, & DAC
            string RNCC = string.Empty;
            string RNVC = string.Empty;
            string DAC = string.Empty;
            if (best03Line.Length >= 5)
            {
                RNCC = best03Line.Substring(3, 1);
                RNVC = best03Line.Substring(4, 1);
                DAC = best03Line.Substring(5, 1);
            }
            // FSC & NIIN
            string FSC = string.Empty;
            string NIIN = string.Empty;
            if (line01.Length >= 15)
            {
                FSC = line01.Substring(2, 4);
                NIIN = line01.Substring(6, 9);
            }
            // NAME, DMIL, INC
            string NAME = string.Empty;
            string DMIL = string.Empty;
            string INC = string.Empty;
            if (line01.Length >= 48)
            {
                NAME = line01.Substring(26, 19).Trim();
                DMIL = line01.Substring(48, 1).Trim();
                INC = line01.Substring(21, 5).Trim();
            }
            if (TEMP1.Substring(0, 12) == "Part Number:")
            {
                TEMP1 += NAME + ")";
            }
            if (TEMP1.Substring(0, 11) == "Part, Cage:")
            {
                TEMP1 += NAME + ")";
                exactMatches++;
            }

            // SL, UI, UI PRICE
            string SL = string.Empty;
            string UI = string.Empty;
            string UIPRICE = string.Empty;
            string QUP = string.Empty;

            if (best05Line.Length >= 24)
            {
                SL = best05Line.Substring(23, 1);
                UI = best05Line.Substring(9, 2);
                UIPRICE = best05Line.Substring(11, 12);
                QUP = best05Line.Substring(8, 1);
            }

            string zeroIndex = "0000" + ouptutIndex.ToString();
            zeroIndex = zeroIndex.Substring(zeroIndex.Length - 4, 4);

            #region PPL fields
            pplColumns[0] = zeroIndex; //INDEX
            pplColumns[1] = ""; //PCCN
            pplColumns[2] = ""; //PLISN
            pplColumns[3] = ""; //INDC
            if (CAGE == "")
            {
                CAGE = cage;
            }
            pplColumns[4] = CAGE;        //CAGE
            if (PN == "")
            {
                PN = pn.ToUpper();
            }
            pplColumns[5] = PN;             //PN
            pplColumns[6] = RNCC;        //RNCC
            pplColumns[7] = RNVC;        //RNVC
            pplColumns[8] = DAC;          //DAC
            pplColumns[9] = ""; //PPSL
            pplColumns[10] = ""; //EC
            if (NAME == "")
            {
                NAME = name.ToUpper();
            }
            pplColumns[11] = NAME;     //NAME
            pplColumns[12] = SL;            //SL
            pplColumns[13] = ""; //SLAC
            pplColumns[14] = ""; //COG
            pplColumns[15] = ""; //MCC
            pplColumns[16] = FSC;           //FSC
            pplColumns[17] = NIIN;          //NIIN
            pplColumns[18] = ""; //NSNSUFF
            pplColumns[19] = ""; //UM
            pplColumns[20] = ""; //UM PRICE
            pplColumns[21] = UI;            //UI
            pplColumns[22] = UIPRICE;  //UI PRICE
            pplColumns[23] = ""; //CONV
            pplColumns[24] = QUP;        //QUP
            pplColumns[25] = ""; //SMR
            pplColumns[26] = DMIL;        //DMIL
            pplColumns[27] = ""; //PLT
            pplColumns[28] = ""; //HCI
            pplColumns[29] = ""; //PSPC
            pplColumns[30] = ""; //PMIC
            pplColumns[31] = ""; //ADPEC
            pplColumns[32] = ""; //NHA
            pplColumns[33] = ""; //ORR
            pplColumns[34] = ""; //QPA
            pplColumns[35] = ""; //QPE
            pplColumns[36] = ""; //MRRI
            pplColumns[37] = ""; //MRRII
            pplColumns[38] = ""; //MRR MOD
            pplColumns[39] = ""; //TQR
            pplColumns[40] = ""; //SAPLISN
            pplColumns[41] = ""; //PRPLISN
            pplColumns[42] = ""; //MAOT
            pplColumns[43] = ""; //MAC
            pplColumns[44] = ""; //NRTS
            pplColumns[45] = ""; //UOC
            pplColumns[46] = ""; //REFDES
            pplColumns[47] = ""; //RDOC
            pplColumns[48] = ""; //RDC
            pplColumns[49] = ""; //SMCC
            pplColumns[50] = ""; //PLCC
            pplColumns[51] = ""; //SMIC
            pplColumns[52] = ""; //AIC
            pplColumns[53] = ""; //AIC QTY
            pplColumns[54] = ""; //MRU
            pplColumns[55] = ""; //RMSS
            pplColumns[56] = ""; //RISS
            pplColumns[57] = ""; //RTLL QTY
            pplColumns[58] = ""; //RSR
            pplColumns[59] = ""; //O-MTD
            pplColumns[60] = ""; //F-MTD
            pplColumns[61] = ""; //H-MTD
            pplColumns[62] = ""; //SRA-MTD
            pplColumns[63] = ""; //D-MTD
            pplColumns[64] = ""; //CED-MTD
            pplColumns[65] = ""; //CAD-MTD
            pplColumns[66] = ""; //O-RCT
            pplColumns[67] = ""; //F-RCT
            pplColumns[68] = ""; //H-RCT
            pplColumns[69] = ""; //SRA-RCT
            pplColumns[70] = ""; //D-RCT
            pplColumns[71] = ""; //CON-RCT
            pplColumns[72] = ""; //O-RTD
            pplColumns[73] = ""; //F-RTD
            pplColumns[74] = ""; //H-RTD
            pplColumns[75] = ""; //SRA-RTD
            pplColumns[76] = ""; //D-RTD
            pplColumns[77] = ""; //DOP1
            pplColumns[78] = ""; //DOP2
            pplColumns[79] = ""; //CTIC
            pplColumns[80] = ""; //AMC
            pplColumns[81] = ""; //AMSC
            pplColumns[82] = ""; //IMC
            pplColumns[83] = ""; //RIP
            pplColumns[84] = ""; //CHANGE AUTHORITY1
            pplColumns[85] = ""; //IC
            pplColumns[86] = ""; //SN FROM
            pplColumns[87] = ""; //SN TO
            pplColumns[88] = ""; //TIC
            pplColumns[89] = ""; //R/S PLISN
            pplColumns[90] = ""; //QTY SHIPPED
            pplColumns[91] = ""; //QTY PROCURED
            pplColumns[92] = ""; //DCN UOC
            pplColumns[93] = ""; //PRORATED ELIN
            pplColumns[94] = ""; //PRORATED QTY
            pplColumns[95] = ""; //LCN
            pplColumns[96] = ""; //ALT LCN
            pplColumns[97] = REMARKS;           //REMARKS
            pplColumns[98] = ""; //TM CODE
            pplColumns[99] = ""; //FIG NUM
            pplColumns[100] = ""; //ITEM NUM
            pplColumns[101] = ""; //TM CHG
            pplColumns[102] = ""; //TM IND
            pplColumns[103] = ""; //QTY FIG
            pplColumns[104] = ""; //WUC/TM FGC
            pplColumns[105] = ""; //BASIS OF ISSUE1
            pplColumns[106] = ""; //BASIS OF ISSUE2
            pplColumns[107] = ""; //CC
            pplColumns[108] = INC;              //INC
            pplColumns[109] = ""; //LRU
            if (PROVNOM == "")
            {
                PROVNOM = ProvNomArray[pnIndex].Trim().ToUpper();
            }
            pplColumns[110] = PROVNOM;   //PROV NOM
            pplColumns[111] = ""; //ALTCAGE1
            pplColumns[112] = ""; //ALTPN1
            pplColumns[113] = ""; //ALTRNCC1
            pplColumns[114] = ""; //ALTRNVC1
            pplColumns[115] = ""; //ALTDAC1
            pplColumns[116] = ""; //ALTPPSL1
            pplColumns[117] = ""; //ALTCAGE2
            pplColumns[118] = ""; //ALTPN2
            pplColumns[119] = ""; //ALTRNCC2
            pplColumns[120] = ""; //ALTRNVC2
            pplColumns[121] = ""; //ALTDAC2
            pplColumns[122] = ""; //ALTPPSL2
            pplColumns[123] = ""; //ALTCAGE3
            pplColumns[124] = ""; //ALTPN3
            pplColumns[125] = ""; //ALTRNCC3
            pplColumns[126] = ""; //ALTRNVC3
            pplColumns[127] = ""; //ALTDAC3
            pplColumns[128] = ""; //ALTPPSL3
            pplColumns[129] = ""; //ALTCAGE4
            pplColumns[130] = ""; //ALTPN4
            pplColumns[131] = ""; //ALTRNCC4
            pplColumns[132] = ""; //ALTRNVC4
            pplColumns[133] = ""; //ALTDAC4
            pplColumns[134] = ""; //ALTPPSL4
            pplColumns[135] = ""; //ALTCAGE5
            pplColumns[136] = ""; //ALTPN5
            pplColumns[137] = ""; //ALTRNCC5
            pplColumns[138] = ""; //ALTRNVC5
            pplColumns[139] = ""; //ALTDAC5
            pplColumns[140] = ""; //ALTPPSL5
            pplColumns[141] = ""; //ALTCAGE6
            pplColumns[142] = ""; //ALTPN6
            pplColumns[143] = ""; //ALTRNCC6
            pplColumns[144] = ""; //ALTRNVC6
            pplColumns[145] = ""; //ALTDAC6
            pplColumns[146] = ""; //ALTPPSL6
            pplColumns[147] = ""; //MATERIAL1
            pplColumns[148] = ""; //MATERIAL2
            pplColumns[149] = ""; //MATERIAL3
            pplColumns[150] = ""; //MATERIAL4
            pplColumns[151] = ""; //RBD
            pplColumns[152] = ""; //SUPP NOMEN 1
            pplColumns[153] = ""; //AEL-A
            pplColumns[154] = ""; //AEL-B
            pplColumns[155] = ""; //AEL-C
            pplColumns[156] = ""; //AEL-D
            pplColumns[157] = ""; //AEL-E
            pplColumns[158] = ""; //AEL-F
            pplColumns[159] = ""; //AEL-G
            pplColumns[160] = ""; //AEL-H
            pplColumns[161] = ""; //SUPP NOMEN 2
            pplColumns[162] = ""; //AFC2
            pplColumns[163] = ""; //AFC QTY2
            pplColumns[164] = ""; //ANC2
            pplColumns[165] = ""; //AOC2
            pplColumns[166] = ""; //AOC QTY2
            pplColumns[167] = ""; //LLTIL1
            pplColumns[168] = ""; //PPL1
            pplColumns[169] = ""; //SFPPL1
            pplColumns[170] = ""; //CBIL1
            pplColumns[171] = ""; //RIL1
            pplColumns[172] = ""; //ISIL1
            pplColumns[173] = ""; //PCL1
            pplColumns[174] = ""; //TTEL1
            pplColumns[175] = ""; //SCPL1
            pplColumns[176] = ""; //DCN1
            pplColumns[177] = ""; //ARF
            pplColumns[178] = ""; //LLTIL2
            pplColumns[179] = ""; //PPL2
            pplColumns[180] = ""; //SFPPL2
            pplColumns[181] = ""; //CBIL2
            pplColumns[182] = ""; //RIL2
            pplColumns[183] = ""; //ISIL2
            pplColumns[184] = ""; //PCL2
            pplColumns[185] = ""; //TTEL2
            pplColumns[186] = ""; //SCPL2
            pplColumns[187] = ""; //DCN2
            pplColumns[188] = ""; //ACC CODE
            pplColumns[189] = ""; //ALT NIIN REL
            pplColumns[190] = ""; //ALT NIIN
            pplColumns[191] = ""; //ALT NIIN REL2
            pplColumns[192] = ""; //ALT NIIN2
            pplColumns[193] = ""; //REFDES2
            pplColumns[194] = ""; //RDOC2
            pplColumns[195] = ""; //CHANGE AUTHORITY2
            pplColumns[196] = ""; //IC2
            pplColumns[197] = ""; //SN FROM2
            pplColumns[198] = ""; //SN TO2
            pplColumns[199] = ""; //TIC2
            pplColumns[200] = ""; //R/S PLISN2
            pplColumns[201] = ""; //QTY SHIPPED2
            pplColumns[202] = ""; //QTY PROCURED2
            pplColumns[203] = ""; //DCN UOC2
            pplColumns[204] = ""; //PRORATED ELIN2
            pplColumns[205] = ""; //PRORATED QTY2
            pplColumns[206] = ""; //LCN2
            pplColumns[207] = ""; //ALT LCN2
            pplColumns[208] = ""; //LENGTH
            pplColumns[209] = ""; //WIDTH
            pplColumns[210] = ""; //HEIGHT
            pplColumns[211] = ""; //WEIGHT
            pplColumns[212] = TEMP1;          //temp1
            #endregion

            return pplColumns;
        }

        private List<string[]> PickBestPart(List<string[]> matchingParts, string pn, string cage)
        {
            List<string[]> bestPart = new List<string[]>() { };
            List<string[]> unsortedParts = matchingParts;

            // if there's only one matching part no more filtering can be done
            if (matchingParts.Count == 1)
            {
                bestPart = matchingParts;
            }
            else
            {
                // if there's more than one match, find the part(s) w/ the best rnvc/rncc combo and discard the rest.
                unsortedParts = ReturnBestRnvcRncc(unsortedParts);
                // if after filtering rnvc/rncc there's only one matching part no more filtering can be done
                if (unsortedParts.Count == 1)
                {
                    bestPart = unsortedParts;
                }
                else
                {
                    // if there's still more than one match, find the part(s) w/ the best agency codes and discard the rest.
                    unsortedParts = ReturnBestAgencyCode(unsortedParts);
                    // if after filtering agency codes there's only one matching part no more filtering can be done
                    if (unsortedParts.Count == 1)
                    {
                        bestPart = unsortedParts;
                    }
                    else
                    {
                        unsortedParts = ReturnBestAgencyCode(unsortedParts);

                    }
                }
            }

            return bestPart;
        }

        private List<string[]> ReturnBestAgencyCode(List<string[]> unsortedParts)
        {
            List<string[]> bestAgencyCode = new List<string[]>() { };
            List<string[]> AgencyCodeDA = new List<string[]>() { };
            List<string[]> AgencyCodeDS = new List<string[]>() { };
            List<string[]> AgencyCodeDF = new List<string[]>() { };
            List<string[]> AgencyCodeDN = new List<string[]>() { };
            List<string[]> AgencyCodeDM = new List<string[]>() { };

            foreach (string[] arr in unsortedParts)
            {
                for (int i = 0; i < arr.Length; i++)
                {
                    if (arr[i].Substring(0, 2) == "05" && arr[i].Substring(2, 2) == "DA")
                    {
                        if (!AgencyCodeDA.Contains(arr))
                        {
                            AgencyCodeDA.Add(arr);
                        }
                    }
                    else if (arr[i].Substring(0, 2) == "05" && arr[i].Substring(2, 2) == "DS")
                    {
                        if (!AgencyCodeDS.Contains(arr))
                        {
                            AgencyCodeDS.Add(arr);
                        }
                    }
                    else if (arr[i].Substring(0, 2) == "05" && arr[i].Substring(2, 2) == "DF")
                    {
                        if (!AgencyCodeDF.Contains(arr))
                        {
                            AgencyCodeDF.Add(arr);
                        }
                    }
                    else if (arr[i].Substring(0, 2) == "05" && arr[i].Substring(2, 2) == "DN")
                    {
                        if (!AgencyCodeDN.Contains(arr))
                        {
                            AgencyCodeDN.Add(arr);
                        }
                    }
                    else if (arr[i].Substring(0, 2) == "05" && arr[i].Substring(2, 2) == "DM")
                    {
                        if (!AgencyCodeDM.Contains(arr))
                        {
                            AgencyCodeDM.Add(arr);
                        }
                    }
                }
            }

            if (AgencyCodeDA.Count != 0)
            {
                bestAgencyCode = AgencyCodeDA;
            }
            else if (AgencyCodeDS.Count != 0)
            {
                bestAgencyCode = AgencyCodeDS;
            }
            else if (AgencyCodeDF.Count != 0)
            {
                bestAgencyCode = AgencyCodeDF;
            }
            else if (AgencyCodeDN.Count != 0)
            {
                bestAgencyCode = AgencyCodeDN;
            }
            else if (AgencyCodeDM.Count != 0)
            {
                bestAgencyCode = AgencyCodeDM;
            }
            else
            {
                bestAgencyCode = unsortedParts;
            }
            return bestAgencyCode;
        }

        private List<string[]> ReturnBestRnvcRncc(List<string[]> unsortedParts)
        {
            List<string[]> bestRnccRnvc = new List<string[]>() { };
            List<string[]> rnccRnvc22 = new List<string[]>() { };
            List<string[]> rnccRnvc32 = new List<string[]>() { };
            List<string[]> rnccRnvc33 = new List<string[]>() { };
            List<string[]> rnccRnvc52 = new List<string[]>() { };
            List<string[]> rnccRnvc53 = new List<string[]>() { };
            List<string[]> rnccRnvc58 = new List<string[]>() { };
            List<string[]> rnccRnvc82 = new List<string[]>() { };

            foreach (string[] arr in unsortedParts)
            {
                for (int i = 0; i < arr.Length; i++)
                {
                    if (arr[i].Substring(0, 2) == "03" && arr[i].Substring(3, 2) == "22")
                    {
                        if (!rnccRnvc22.Contains(arr))
                        {
                            rnccRnvc22.Add(arr);
                        }
                    }
                    else if (arr[i].Substring(0, 2) == "03" && arr[i].Substring(3, 2) == "32")
                    {
                        if (!rnccRnvc32.Contains(arr))
                        {
                            rnccRnvc32.Add(arr);
                        }
                    }
                    else if (arr[i].Substring(0, 2) == "03" && arr[i].Substring(3, 2) == "33")
                    {
                        if (!rnccRnvc33.Contains(arr))
                        {
                            rnccRnvc33.Add(arr);
                        }
                    }
                    else if (arr[i].Substring(0, 2) == "03" && arr[i].Substring(3, 2) == "52")
                    {
                        if (!rnccRnvc52.Contains(arr))
                        {
                            rnccRnvc52.Add(arr);
                        }
                    }
                    else if (arr[i].Substring(0, 2) == "03" && arr[i].Substring(3, 2) == "53")
                    {
                        if (!rnccRnvc53.Contains(arr))
                        {
                            rnccRnvc53.Add(arr);
                        }
                    }
                    else if (arr[i].Substring(0, 2) == "03" && arr[i].Substring(3, 2) == "82")
                    {
                        if (!rnccRnvc82.Contains(arr))
                        {
                            rnccRnvc82.Add(arr);
                        }
                    }
                }
            }

            if (rnccRnvc22.Count != 0)
            {
                bestRnccRnvc = rnccRnvc22;
            }
            else if (rnccRnvc32.Count != 0)
            {
                bestRnccRnvc = rnccRnvc32;
            }
            else if (rnccRnvc33.Count != 0)
            {
                bestRnccRnvc = rnccRnvc33;
            }
            else if (rnccRnvc52.Count != 0)
            {
                bestRnccRnvc = rnccRnvc52;
            }
            else if (rnccRnvc53.Count != 0)
            {
                bestRnccRnvc = rnccRnvc53;
            }
            else if (rnccRnvc82.Count != 0)
            {
                bestRnccRnvc = rnccRnvc82;
            }
            else
            {
                bestRnccRnvc = unsortedParts;
            }

            return bestRnccRnvc;
        }

        private string Sort05Lines(List<string> unsorted05Lines)
        {
            string best05Line = string.Empty;
            if (unsorted05Lines.Count == 1)
            {
                foreach (string line in unsorted05Lines)
                {
                    best05Line = line;
                }
            }
            else
            {
                List<string> agencyCodeDA = new List<string>() { };
                List<string> agencyCodeDS = new List<string>() { };
                List<string> agencyCodeDF = new List<string>() { };
                List<string> agencyCodeDN = new List<string>() { };
                List<string> agencyCodeDM = new List<string>() { };
                List<string> sorted05Lines = new List<string>() { };
                List<string> other = new List<string>() { };

                foreach (string line in unsorted05Lines)
                {
                    if (line.Substring(2, 2).ToUpper() == "DA")
                    {
                        agencyCodeDA.Add(line);
                    }
                    else if (line.Substring(2, 2).ToUpper() == "DS")
                    {
                        agencyCodeDS.Add(line);
                    }
                    else if (line.Substring(2, 2).ToUpper() == "DF")
                    {
                        agencyCodeDF.Add(line);
                    }
                    else if (line.Substring(2, 2).ToUpper() == "DN")
                    {
                        agencyCodeDN.Add(line);
                    }
                    else if (line.Substring(2, 2).ToUpper() == "DM")
                    {
                        agencyCodeDM.Add(line);
                    }
                    else
                    {
                        other.Add(line);
                    }
                }

                foreach (string line in agencyCodeDA)
                {
                    sorted05Lines.Add(line);
                }
                foreach (string line in agencyCodeDS)
                {
                    sorted05Lines.Add(line);
                }
                foreach (string line in agencyCodeDF)
                {
                    sorted05Lines.Add(line);
                }
                foreach (string line in agencyCodeDN)
                {
                    sorted05Lines.Add(line);
                }
                foreach (string line in agencyCodeDM)
                {
                    sorted05Lines.Add(line);
                }
                foreach (string line in sorted05Lines)
                {
                    best05Line = line;
                    break;
                }

            }
            return best05Line;
        }

        private string Sort03Lines(List<string> unsorted03Lines)
        {
            string best03Line = string.Empty;
            List<string> list22 = new List<string>() { };
            List<string> list32 = new List<string>() { };
            List<string> list33 = new List<string>() { };
            List<string> list52 = new List<string>() { };
            List<string> list53 = new List<string>() { };
            List<string> list82 = new List<string>() { };
            List<string> other = new List<string>() { };
            List<string> sorted03Lines = new List<string>() { };
            foreach (var line03 in unsorted03Lines)
            {
                if (line03.Substring(3, 2) == "22")
                {
                    list22.Add(line03);
                }
                else if (line03.Substring(3, 2) == "32")
                {
                    list32.Add(line03);
                }
                else if (line03.Substring(3, 2) == "33")
                {
                    list33.Add(line03);
                }
                else if (line03.Substring(3, 2) == "52")
                {
                    list52.Add(line03);
                }
                else if (line03.Substring(3, 2) == "53")
                {
                    list53.Add(line03);
                }
                else if (line03.Substring(3, 2) == "82")
                {
                    list82.Add(line03);
                }
                else
                {
                    other.Add(line03);
                }
            }

            if (list22.Count > 0)
            {
                foreach (var line in list22)
                {
                    sorted03Lines.Add(line);
                }
            }
            if (list32.Count > 0)
            {
                foreach (var line in list32)
                {
                    sorted03Lines.Add(line);
                }
            }
            if (list33.Count > 0)
            {
                foreach (var line in list33)
                {
                    sorted03Lines.Add(line);
                }
            }
            if (list52.Count > 0)
            {
                foreach (var line in list52)
                {
                    sorted03Lines.Add(line);
                }
            }
            if (list53.Count > 0)
            {
                foreach (var line in list53)
                {
                    sorted03Lines.Add(line);
                }
            }
            if (list82.Count > 0)
            {
                foreach (var line in list82)
                {
                    sorted03Lines.Add(line);
                }
            }

            if (sorted03Lines.Count >= 1)
            {
                foreach (var line in sorted03Lines)
                {
                    if (best03Line == "")
                    {
                        best03Line = line;
                    }
                }
            }

            return best03Line;
        }

        public string StripPunctuation(string s)
        {
            var sb = new StringBuilder();
            foreach (char c in s)
            {
                if (!char.IsPunctuation(c))
                    sb.Append(c);
            }
            return sb.ToString();
        }

        private void ExportToExcel_Click(object sender, EventArgs e)
        {
            ExportToExcel(table);
        }

        private void ExportToExcel(DataTable dt)
        {
            SaveFileDialog sfd = new SaveFileDialog();
            sfd.Filter = "Excel file (*.xlsx)|*.xlsx";

            if (sfd.ShowDialog() == DialogResult.OK)
            {
                using (ExcelPackage pck = new ExcelPackage())
                {
                    ExcelWorksheet ws = pck.Workbook.Worksheets.Add("PPL Data");
                    ws.Cells["A1"].LoadFromDataTable(dt, true);
                    pck.SaveAs(new FileInfo(sfd.FileName));
                }
            }

            MessageBox.Show("Data exported successfully.\n\n" + sfd.FileName.ToString());
        }

        private void Load036_Click(object sender, EventArgs e)
        {
            Load036PartsData();
        }

        private void Load036PartsData()
        {
            OpenFileDialog dataFile036 = new OpenFileDialog();
            dataFile036.Filter = "Text Files|*.txt|036 Files|*.036";
            dataFile036.Title = "Select .036 Data File";
            dataFile036.Multiselect = true;
            if (dataFile036.ShowDialog() == System.Windows.Forms.DialogResult.OK)
            {
                List<string[]> DataFromFiles = new List<string[]>() { };

                string filesPicked = string.Empty;
                foreach (string file036 in dataFile036.FileNames)
                {
                    string[] tempStringArray = File.ReadAllLines(file036);
                    DataFromFiles.Add(tempStringArray);
                }

                List<string[]> all036Parts = Parse036Parts(DataFromFiles);


            }
        }

        private List<string[]> Parse036Parts(List<string[]> dataFromFiles)
        {
            List<string> CombinedDataFromFilesList = new List<string>() { };
            List<string[]> AllPartsFromFiles = new List<string[]>() { };
            foreach (string[] fileData in dataFromFiles)
            {
                for (int i = 0; i < fileData.Length; i++)
                {
                    if (fileData[i].Trim() != "")
                    {
                        CombinedDataFromFilesList.Add(fileData[i]);
                    }                    
                }
            }

            List<string> TempListToArray = new List<string>() { };
            bool partFound = false;
            int CombinedDataIndex = 0;
            foreach (string datarow in CombinedDataFromFilesList)
            {
                CombinedDataIndex++;
                if (!partFound)
                {
                    if (GetLineType(datarow) == "01A")
                    {
                        partFound = true;
                        TempListToArray.Add(datarow);
                    }
                }
                else
                {
                    if (CombinedDataIndex != CombinedDataFromFilesList.Count)
                    {
                        if (GetLineType(datarow) == "01A")
                        {
                            AllPartsFromFiles.Add(TempListToArray.ToArray());
                            TempListToArray.Clear();
                            TempListToArray.Add(datarow);
                        }
                        else
                        {
                            TempListToArray.Add(datarow);
                        }
                    }
                    else
                    {
                        TempListToArray.Add(datarow);
                        AllPartsFromFiles.Add(TempListToArray.ToArray());
                    }
                }
            }

            return AllPartsFromFiles;
        }

        private string GetLineType(string datarow)
        {
            string lineType = string.Empty;
            if (datarow.Trim() != "")
            {
                if (datarow.Length >= 3)
                {
                    lineType = datarow.Substring(datarow.Length-3,3);
                }                
            }
            return lineType;
        }
    }
}

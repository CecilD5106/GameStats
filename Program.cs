//Use team statistics to create game statistics
using Microsoft.Office.Interop.Excel;
using System;
using System.Collections.Generic;

namespace GameStats
{
    class Program
    {
        static void Main(string[] args)
        {
            //Set Variables
            string path = "E:\\Code\\VSCode\\Node\\CFB01\\2019CFPickem.xlsx";
            Application excel = new Application();
            Workbook wb = excel.Workbooks.Open(path);
            try
            {
                List<string> sTabs = new List<string>(new string[] { "Air Force", "Akron", "Alabama", "Appalachian State", "Arizona",
                "Arizona State", "Arkansas", "Arkansas State", "Army", "Auburn", "Ball State", "Baylor", "Boise State", "Boston College",
                "Bowling Green", "Buffalo", "BYU", "California", "Central Michigan", "Charlotte", "Cincinnati", "Clemson", "Coastal Carolina",
                "Colorado", "Colorado State", "UConn", "Duke", "East Carolina", "Eastern Michigan", "Fresno State", "Florida", "Florida Atlantic",
                "Florida International", "Florida State", "Georgia", "Georgia State", "Georgia Tech", "Georgia Southern", "Hawai'i",
                "Houston", "Illinois", "Indiana", "Iowa", "Iowa State", "Kansas", "Kansas State", "Kent State", "Kentucky", "Liberty",
                "Louisiana", "UL Monroe", "Louisiana Tech", "Louisville", "LSU", "Marshall", "Maryland", "Memphis", "Miami", "Miami (OH)",
                "Michigan", "Michigan State", "Middle Tennessee", "Minnesota", "Mississippi State", "Missouri", "Navy", "NC State", "Nebraska",
                "Nevada", "New Mexico", "New Mexico State", "North Carolina", "North Texas", "Northern Illinois", "Northwestern", "Notre Dame",
                "Ohio", "Ohio State", "Oklahoma", "Oklahoma State", "Old Dominion", "Ole Miss", "Oregon", "Oregon State", "Penn State",
                "Pittsburgh", "Purdue", "Rice", "Rutgers", "San Diego State", "San Jose State", "SMU", "South Alabama", "South Carolina",
                "South Florida", "Southern Mississippi", "Stanford", "Syracuse", "TCU", "Temple", "Tennessee", "Texas", "Texas A&M", "Texas State",
                "Texas Tech", "Toledo", "Troy", "Tulane", "Tulsa", "UAB", "UCF", "UCLA", "UMass", "UNLV", "USC", "Utah", "Utah State", "UTEP", "UTSA",
                "Vanderbilt", "Virginia", "Virginia Tech", "Wake Forest", "Washington", "Washington State", "West Virginia", "Western Kentucky",
                "Western Michigan", "Wisconsin", "Wyoming" });

                //Loop through the teams in tabs
                foreach (var sTab in sTabs)
                {
                    Worksheet ws = wb.Worksheets[sTab];
                    if (ws.Cells[7, 4].Value != 0)
                    {
                        //Get blank row
                        int i = 12;
                        while (ws.Cells[i, 1].Value != null)
                        {
                            i++;
                        }

                        //Put games stats in blank row
                        ws.Cells[i, 1].Value = ws.Cells[7, 2].Value;
                        ws.Cells[i, 2].Value = ws.Cells[7, 4].Value;
                        ws.Cells[i, 3].Value = ws.Cells[7, 6].Value;
                        ws.Cells[i, 4].Value = ws.Cells[7, 8].Value;
                        ws.Cells[i, 5].Value = ws.Cells[7, 10].Value;
                        ws.Cells[i, 6].Value = ws.Cells[7, 12].Value;
                        ws.Cells[i, 7].Value = ws.Cells[7, 14].Value;
                        ws.Cells[i, 8].Value = ws.Cells[7, 16].Value;
                        //Determine if game was a win or a loss
                        if (ws.Cells[i, 1].Value > ws.Cells[i, 5].Value)
                        {
                            ws.Cells[i, 9].Value = 1;
                        }
                        else
                        {
                            ws.Cells[i, 9].Value = 0;
                        }
                    }
                }

                wb.Save();
                excel.Quit();
            }
            catch (Exception e)
            {
                excel.Quit();
                Console.WriteLine(e.ToString());
                throw;
            }
            finally
            {
                excel.Quit();
            }
        }
    }
}

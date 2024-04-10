using Microsoft.Office.Interop.Excel;
using System;
using System.Collections.Generic;
using System.Data;
using Excel = Microsoft.Office.Interop.Excel;

namespace Application
{
    internal class Piece
    {
        Dictionary<String, List<Data.Data>> measureValues;
        String currentMeasureType;

        public Piece() 
        {
            this.measureValues = new Dictionary<string, List<Data.Data>>();
            this.currentMeasureType = "";
        }

        public int GetLinesToWriteNumber()
        {
            int lineNb = 0;

            foreach(List<Data.Data> data in measureValues.Values) 
            { 
                lineNb++;

                lineNb += data.Count;
            }

            return lineNb;
        }

        public void AddMeasureType(String measureType)
        {
            this.currentMeasureType = measureType;

            this.measureValues.Add(this.currentMeasureType, new List<Data.Data>());
        }

        public void AddData(Data.Data data)
        {
            this.measureValues[this.currentMeasureType].Add(data);
        }

        public void SetValues(List<double> values)
        {
            this.measureValues[this.currentMeasureType].Last().SetValues(values);
        }

        public void WriteBaseValues(Excel.Workbook wb, int line, int col)
        {
            /*for(int i = 0; i < pieceData.Count; i++) 
            {
                wb.ActiveSheet.Cells[line, col].Value = this.measureTypes[i];
                line++;

                for(int j = 0; j < pieceData[i].Count; j++)
                {
                    wb.ActiveSheet.Cells[line, col].Value = this.pieceData[i][j].GetNominalValue();
                    col += 2;
                    wb.ActiveSheet.Cells[line, col].Value = this.pieceData[i][j].GetTolPlus();
                    col++;
                    wb.ActiveSheet.Cells[line, col].Value = this.pieceData[i][j].GetTolMinus();
                    line++;
                    col -= 3;
                }
            }*/
        }

        public void WriteValues(Excel.Workbook wb, int line, int col)
        {
            Excel.Worksheet ws = wb.Sheets["Mesures"];
            int linesWritten = 0;
            int pageNumber = 1;

            foreach(var item in this.measureValues)
            {
                col++;
                ws.Cells[line, col].Value = item.Key;
                line++;
                linesWritten++;
                col--;

                if(linesWritten == 22)
                {
                    pageNumber++;

                    ws = wb.Sheets["Mesures (" + pageNumber.ToString() + ")"];

                    line -= linesWritten;
                    linesWritten = 0;
                }

                foreach (Data.Data data in item.Value)
                {
                    col++;
                    col++;
                    ws.Cells[line, col].Value = data.GetNominalValue();
                    col++;
                    col++;
                    ws.Cells[line, col].Value = data.GetTolPlus();
                    col++;
                    ws.Cells[line, col].Value = data.GetTolMinus();
                    col++;
                    ws.Cells[line, col].Value = data.GetValue();

                    line++;
                    linesWritten++;
                    col -= 6;

                    if (linesWritten == 22)
                    {
                        pageNumber++;
                        
                        ws = wb.Sheets["Mesures (" + pageNumber.ToString() + ")"];

                        line -= linesWritten;
                        linesWritten = 0;
                    }
                }
            }
        }
    }
}

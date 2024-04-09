using Microsoft.Office.Interop.Excel;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Reflection;
using System.Text;
using System.Threading.Tasks;
using Excel = Microsoft.Office.Interop.Excel;

namespace Application
{
    internal class Piece
    {
        private List<List<Data.Data>> pieceData;
        private List<String> measureTypes;

        public Piece() 
        {
            this.pieceData = new List<List<Data.Data>>();
            this.measureTypes = new List<String>();
        }

        public int GetLinesToWriteNumber()
        {
            int lineNb = 0;

            for(int i = 0; i < this.pieceData.Count; i++) 
            {
                lineNb++;

                lineNb += this.pieceData[i].Count;
            }

            return lineNb;
        }

        public void AddMeasureType(String measureType)
        {
            this.measureTypes.Add(measureType);
            this.pieceData.Add(new List<Data.Data>());
        }

        public void AddData(Data.Data data)
        {
            this.pieceData[pieceData.Count - 1].Add(data);
        }

        public void SetValues(List<double> values)
        {
            int i = pieceData.Count - 1;
            int j = this.pieceData[i].Count - 1;

            this.pieceData[i][j].SetValues(values);
        }

        public void WriteBaseValues(Excel.Workbook wb, int line, int col)
        {
            for(int i = 0; i < pieceData.Count; i++) 
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
            }
        }

        public void WriteValues(Excel.Workbook wb, int line, int col)
        {
            Excel.Worksheet ws = wb.Sheets["Mesures"];
            int linesWritten = 0;
            int pageNumber = 1;

            for(int i = 0; i < pieceData.Count; i++)
            {
                col++;
                ws.Cells[line, col].Value = this.measureTypes[i];
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

                for (int j = 0; j < this.pieceData[i].Count; j++)
                {
                    col++;
                    col++;
                    ws.Cells[line, col].Value = this.pieceData[i][j].GetNominalValue();
                    col++;
                    col++;
                    ws.Cells[line, col].Value = this.pieceData[i][j].GetTolPlus();
                    col++;
                    ws.Cells[line, col].Value = this.pieceData[i][j].GetTolMinus();
                    col++;
                    ws.Cells[line, col].Value = this.pieceData[i][j].GetValue();

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

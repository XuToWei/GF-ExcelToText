//------------------------------------------------------------
// ExcelToTxt
// Copyright Xu wei
//------------------------------------------------------------
using System.IO;
using System.Collections;
using System.Collections.Generic;
using System.Text;
using System.Text.RegularExpressions;
using UnityEngine;
using NPOI.SS.UserModel;
using NPOI.XSSF.UserModel;
using UnityEditor;
using UnityGameFramework.Runtime;

namespace GameMain.Editor.ExcelTool
{
    public sealed class ExcelToTxt
    {
        private static readonly Regex NameRegex = new Regex(@"^[A-Z][A-Za-z0-9_]*$");
        
        public static void Convert()
        {
            string excelsFloder = $"{Application.dataPath}/../Excels/";
            if (!Directory.Exists(excelsFloder))
            {
                Log.Error("{0} is not exist!", excelsFloder);
                return;
            }

            string[] excelFiles = Directory.GetFiles(excelsFloder);
            foreach (var excelFile in excelFiles)
            {
                if (!excelFile.EndsWith(".xlsx") || excelFile.Contains("~$"))
                    continue;
                FileStream fileStream = new FileStream(excelFile, FileMode.Open, FileAccess.Read);
                IWorkbook workbook = new XSSFWorkbook(fileStream);
                for (int s = 0; s < workbook.NumberOfSheets; s++)
                {
                    ISheet sheet = workbook.GetSheetAt(s);
                    if (sheet.LastRowNum < 1)
                        continue;
                    if (!string.Equals(sheet.GetRow(0).Cells[0].ToString(), "#"))
                        continue;
                    string fileName = sheet.GetRow(0).Cells[1].ToString();
                    if (string.IsNullOrWhiteSpace(fileName))
                    {
                        Debug.LogErrorFormat("{0} has not datable name!", fileName);
                        continue;
                    }
                    if (!NameRegex.IsMatch(fileName))
                    {
                        Debug.LogErrorFormat("{0} has wrong datable name!", fileName);
                        continue;
                    }

                    string fileFullPath = $"{Application.dataPath}/GameMain/DataTables/{fileName}.txt";
                    if (File.Exists(fileFullPath))
                    {
                        File.Delete(fileFullPath);
                    }

                    List<string> sContents = new List<string>();
                    StringBuilder sb = new StringBuilder();
                    if (sheet.LastRowNum < 3)
                    {
                        Debug.LogErrorFormat("{0} has wrong row num!", fileFullPath);
                        continue;
                    }

                    IRow row1 = sheet.GetRow(1);
                    int columnCount = 1;
                    for (int i = 0; i < row1.Cells.Count; i++)
                    {
                        if (string.IsNullOrWhiteSpace(row1.Cells[i].ToString()))
                            continue;
                        columnCount++;
                    }

                    for (int i = 0; i <= sheet.LastRowNum + 1; i++)
                    {
                        sb.Clear();
                        IRow row = sheet.GetRow(i);
                        if (row == null || row.Cells == null)
                        {
                            continue;
                        }

                        bool needContinue = true;
                        foreach (var cell in row.Cells)
                        {
                            if (cell != null && !string.IsNullOrWhiteSpace(cell.ToString()))
                                needContinue = false;
                        }

                        if (needContinue)
                        {
                            continue;
                        }

                        int ci = 0;
                        for (int j = 0; j < columnCount; j++)
                        {
                            if (ci >= row.Cells.Count)
                            {
                                sb.Append("");
                            }
                            else
                            {
                                ICell cell = row.Cells[ci];
                                if (cell.ColumnIndex > j)
                                {
                                    sb.Append("");
                                }
                                else
                                {
                                    sb.Append(cell);
                                    ci++;
                                }
                            }

                            if (j != columnCount - 1)
                            {
                                sb.Append('\t');
                            }
                        }

                        sContents.Add(sb.ToString());
                    }

                    File.WriteAllLines(fileFullPath, sContents, Encoding.UTF8);
                    Debug.LogFormat("更新Excel表格：{0}", fileFullPath);
                }
            }

            AssetDatabase.Refresh();
        }
    }
}

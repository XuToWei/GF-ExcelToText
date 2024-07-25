//------------------------------------------------------------
// ExcelToTxt
// Copyright Xu wei
//------------------------------------------------------------
using System.IO;
using System.Collections;
using System.Collections.Generic;
using System.Text;
using System.Text.RegularExpressions;
using GameFramework;
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
        
        static string excelsFloder = $"{Application.dataPath}/../Excels/";
        
        public static void Convert()
        {
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
                                    //处理公式
                                    if (cell.CellType == CellType.Formula)
                                    {
                                        cell.SetCellType(CellType.String);
                                    }
                                    string cellString = cell.ToString();
                                    //处理excel中的换行符
                                    while (cellString.IndexOf('\n') >= 0)
                                    {
                                        cellString = cellString.Remove(cellString.IndexOf('\n'), 1);
                                    }
                                    sb.Append(cellString);
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

            ConfigExcelToTxt();
            AssetDatabase.Refresh();
        }
        
        private static void ConfigExcelToTxt()
        {
            if (!Directory.Exists(excelsFloder))
            {
                Debug.LogError($"{excelsFloder} is not exist!");
                return;
            }
            List<string> excelFiles = new List<string>(Directory.GetFiles(excelsFloder));
            Dictionary<string, string> cachedKeyDict = new Dictionary<string, string>();
            var paths = Directory.GetDirectories(excelsFloder);
            foreach (var item in paths)
            {
                excelFiles.AddRange(Directory.GetFiles(item));
            }
            foreach (var excelName in excelFiles)
            {
                string excelFile = Utility.Path.GetRegularPath(excelName);
                if (!excelFile.EndsWith(".xlsx") || excelFile.Contains("~$") || excelFile.Contains("~"))
                    continue;
                FileStream fileStream = new FileStream(excelFile, FileMode.Open, FileAccess.Read);
                IWorkbook workbook = new XSSFWorkbook(fileStream);
                fileStream.Close();
                for (int s = 0; s < workbook.NumberOfSheets; s++)
                {
                    ISheet sheet = workbook.GetSheetAt(s);
                    if (sheet.LastRowNum < 1)
                        continue;
                    if(sheet.GetRow(0) == null)
                        continue;
                    if(sheet.GetRow(0).Cells.Count < 2)
                        continue;
                    if (!string.Equals(sheet.GetRow(0).Cells[0].ToString(), "#"))
                        continue;
                    string fileName = sheet.GetRow(0).Cells[1].ToString();
                    if(!fileName.EndsWith("Config"))
                        continue;
                    if (string.IsNullOrWhiteSpace(fileName))
                    {
                        Debug.LogErrorFormat("{0} has not config name!", fileName);
                        continue;
                    }
                    if (!NameRegex.IsMatch(fileName))
                    {
                        Debug.LogErrorFormat("{0} has wrong config name!", fileName);
                        continue;
                    }

                    string fileFullPath = $"Assets/GameMain/Configs/{fileName}.txt";
                    if (File.Exists(fileFullPath))
                    {
                        File.Delete(fileFullPath);
                    }

                    List<string> sContents = new List<string>();
                    StringBuilder sb = new StringBuilder();
                    if (sheet.LastRowNum < 2)
                    {
                        Debug.LogErrorFormat("{0} has wrong row num!", fileFullPath);
                        continue;
                    }

                    IRow row1 = sheet.GetRow(1);
                    int columnCount = 4;

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
                            ICell cell = null;
                            if (ci >= row.Cells.Count)
                            {
                                sb.Append("");
                            }
                            else
                            {
                                cell = row.Cells[ci];
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
                            
                            if ( i >= 4 && j == 1 && !excelFile.Contains("data") && !excelFile.Contains("res"))
                            {
                                string sCell = cell.ToString();
                                if (cell != null && !string.IsNullOrEmpty(sCell))
                                {
                                    if (cachedKeyDict.ContainsKey(sCell))
                                    {
                                        throw new GameFrameworkException(Utility.Text.Format("Config1:{0} , Config2:{1}, key: {2} is repeated!", excelFile, cachedKeyDict[sCell], sCell));
                                    }
                                    cachedKeyDict.Add(sCell, excelFile);
                                }
                                else
                                {
                                    throw new GameFrameworkException(Utility.Text.Format("Config:{0} , row: {1} is wrong format!", excelFile, i + 1));
                                }
                            }
                        }

                        sContents.Add(sb.ToString());
                    }

                    File.WriteAllLines(fileFullPath, sContents, Encoding.UTF8);
                    Debug.LogFormat("更新配置Excel表格：{0}", fileFullPath);
                }
            }
            AssetDatabase.Refresh();
        }
    }
}

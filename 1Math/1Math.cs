﻿using AzureCognitiveTranslator;
using System;
using System.Collections;
using System.Collections.Generic;
using System.Net.Http;
using System.Threading;
using System.Threading.Tasks;
using Excel = Microsoft.Office.Interop.Excel;
using System.Text.RegularExpressions;
using System.Diagnostics;
namespace _1Math
{
    public static class Test
    {
    }
    internal static class ExcelStatic
    {
        internal static Excel.Range GetSelectionAsRange()
        {
            if (!(Globals.ThisAddIn.Application.Selection is Excel.Range selection))
            {
                throw new Exception("PleaseSelectExcelRange");
            }
            else
            {
                return selection;
            }
        }
        public static object[,] ToObjects(Excel.Range range)
        {
            if (!CheckContiguous(range))
            {
                throw new Exception("NotContiguousRange");
            }
            object[,] result=(object[,])Array.CreateInstance(typeof(object), new int[2] { range.Rows.Count, range.Columns.Count }, new int[2] { 1, 1 });
            if (CheckSingle(range))
            {
                result[1, 1] = range.Value;
            }
            else
            {
                result = range.Value;
            }
            return result;
        }
        private static bool CheckContiguous(Excel.Range range)
        {
            return range.Areas.Count == 1;
        }
        private static bool CheckSingle(Excel.Range range)
        {
            return range.Count==1;
        }
        public static void StartTask()
        {
            Globals.ThisAddIn.Application.ScreenUpdating = false;
        }
        public static void EndTask()
        {
            Globals.ThisAddIn.Application.ScreenUpdating = true;
        }
        public static int[] ResultOffset { get; set; } = new int[2] { 0, 1 };
    }
    public static class Controller
    {
        public static async Task TranslateSelectionAsync(string toLanguageCode, Translator translator,CancellationToken cancellationToken=new CancellationToken())
        {
            Excel.Range selection = ExcelStatic.GetSelectionAsRange();
            int m = selection.Rows.Count, n = selection.Columns.Count;
            for (int i = 0; i < m; i++)
            {
                for (int j = 0; j < n; j++)
                {
                    translator.AddContent(selection[i + 1, j + 1].Value.ToString());
                }
            }
            List<string> translation;
            translation = await translator.TranslateAsync(toLanguageCode,cancellationToken);
            string[,] translationArr = new string[m, n];
            int t = 0;

            for (int i = 0; i < m; i++)
            {
                for (int j = 0; j < n; j++)
                {
                    translationArr[i, j] = translation[t];
                    t++;
                }
            }
            if (!cancellationToken.IsCancellationRequested)
            {
                selection.Offset[m * ExcelStatic.ResultOffset[0], n * ExcelStatic.ResultOffset[1]].Value = translationArr;
            }
            else
            {
                throw new TaskCanceledException();
            }
        }
    }
    //internal class ExcelConcurrentTask
    //{
    //    ExcelConcurrent _excelConcurrent;
    //    internal ExcelConcurrentTask(ExcelConcurrent excelConcurrent)
    //    {
    //        _excelConcurrent = excelConcurrent;
    //    }
    //    internal async Task StartAsync()
    //    {
    //        CancellationTokenSource cancellationTokenSource = new CancellationTokenSource();
    //        CancellationToken cancellationToken = cancellationTokenSource.Token;
    //        StatusForm statusForm = new StatusForm();
    //        statusForm.Show();
    //        _excelConcurrent.Reportor.MessageChange += statusForm.ChangeMessage;
    //        _excelConcurrent.Reportor.ProgressChange += statusForm.ChangeProgress;
    //        statusForm.FormClosing += (object s, System.Windows.Forms.FormClosingEventArgs e) => cancellationTokenSource.Cancel();
    //        await _excelConcurrent.StartAsync(cancellationToken);
    //    }
    //}
   
   
}

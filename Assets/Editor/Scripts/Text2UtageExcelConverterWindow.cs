using System;
using System.Collections;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text.RegularExpressions;
using NPOI.HSSF.UserModel;
using NPOI.SS.UserModel;
using NPOI.XSSF.UserModel;
using UnityEditor;
using UnityEngine;

namespace Utage {
    public class Text2UtageExcelConverterWindow : EditorWindow {
        public const string MenuToolRoot = "Tools/Utage/";

        readonly string[] columnHedaers = {
            "Command",
            "Arg1",
            "Arg2",
            "Arg3",
            "Arg4",
            "Arg5",
            "Arg6",
            "Text",
            "PageCtrl",
            "Voice",
            "WindowType",
            "English"
        };

        static public Text2UtageExcelConverter Converter {
            get {
                return converter;
            }
            set {
                if (converter != value) {
                    converter = value;
                }
            }
        }
        //プロジェクトデータ
        static Text2UtageExcelConverter converter;

        [MenuItem (MenuToolRoot + "Text2ExcelConverter")]
        static void CreateNewProject () {
            GetWindow (typeof (Text2UtageExcelConverterWindow), false, "Text2ExcelConverter");
        }

        void OnGUI () {
            Text2UtageExcelConverter project = EditorGUILayout.ObjectField ("", Converter, typeof (Text2UtageExcelConverter), false) as Text2UtageExcelConverter;
            if (project != Converter) {
                Converter = project;
            }

            if (GUILayout.Button ("Convert", GUILayout.Width (180))) {
                if (Converter == null) {
                    Debug.Log ("ConverterのScriptableObjectが指定されていません");
                }
                Convert ();
            }
        }

        private void Convert () {
            foreach (var chapter in converter.ExcelChapters) {
                foreach (var text in chapter.TextFiles) {
                    Text2Excel (chapter.ExcelFile, text);
                }
            }
        }

        private void Text2Excel (UnityEngine.Object excelFile, UnityEngine.Object textFile) {

            var excelPath = AssetDatabase.GetAssetPath (excelFile);
            var textPath = AssetDatabase.GetAssetPath (textFile);

            Debug.Log ("Convert :" + textPath);

            IWorkbook writeBook;
            FileStream fs = null;

            // 書込用Excelファイルの読込
            using (fs = File.OpenRead (excelPath)) {
                writeBook = WorkbookFactory.Create (fs);
            }
            var sheetName = Path.GetFileNameWithoutExtension (textPath);
            var targetSheet = writeBook.GetSheet (sheetName);

            if (targetSheet != null) {
                writeBook.RemoveSheetAt (writeBook.GetSheetIndex (sheetName));
            }

            targetSheet = writeBook.CreateSheet (Path.GetFileNameWithoutExtension (textPath));

            var sheet = new Text2Utage.Sheet (targetSheet, Converter.NovelMode);

            using (System.IO.StreamReader sr = new System.IO.StreamReader (textPath)) {
                var textData = sr.ReadToEnd ();
                string[] splitted = textData.Replace ("\r", "").Split ('\n');

                for (int i = 0; i < splitted.Length; i++) {
                    if (splitted[i] == "") continue;
                    //ノベルモードで[p]が存在しない場合、自動でいれる
                    if (converter.NovelMode && !IsNewLineTag (splitted.Skip (i))) {
                        int rowSum = 0;
                        for (int search = i; search < splitted.Length -1; search++) {

                            var nowRow = Mathf.FloorToInt (splitted[search].Length / converter.TextMatrixXY.x) + 1;
                            if(nowRow > converter.TextMatrixXY.y)
                            {
                                splitted[search] = splitted[search] + "[p]";
                                break;
                            }

                            rowSum += nowRow;

                            var nextRow = Mathf.FloorToInt (splitted[search + 1].Length / converter.TextMatrixXY.x) + 1;

                            Debug.Log (splitted[search] + splitted[search].Length + "Row" + rowSum);
                            if ((rowSum + nextRow)> converter.TextMatrixXY.y) {
                                splitted[search] = splitted[search] + "[p]";
 //                               Debug.Log ("index : "+ search + " " +splitted[search]);
                                break;
                            }
                        }
                    }

                    var line = splitted[i];
                    sheet.SetLine (line);
                }

                /*
                foreach (var line in splitted) {


                    if (line == "")
                        continue;

                    var firstChar = line[0];
                    var targetRow = targetSheet.CreateRow (rowCount);

                    if (!isOneByteChar (firstChar)) {
                        var columnIndex = Array.IndexOf (columnHedaers, "Text");
                        var textCell = targetRow.CreateCell (columnIndex);
                        var pageClear = !converter.NovelMode;
                        var setText = line;

                        if (converter.NovelMode) {
                            if (IsNewLineTag (line)) {
                                setText = line.Replace ("[p]", "").Replace ("@p", "");
                            } else {
                                SetPageCtrlInputBr (targetRow);
                            }
                        }
                        textCell.SetCellValue (setText);
                    } else {
                        if (firstChar == ';' || firstChar == '/') {
                            continue;
                        }

                        if (IsNewLineTag (line)) {
                            var beforeRow = targetSheet.GetRow (rowCount - 1);
                            var beforeColumnIndex = Array.IndexOf (columnHedaers, "PageCtrl");
                            var pageCtrlCell = beforeRow.CreateCell (8);
                            pageCtrlCell.SetCellValue ("");
                            continue;
                        }
                        var commandLine = line.Trim ('[', ']', '@').Replace (" =", "=").Replace ("= ", "=");

                        string[] commandSplitted = commandLine.Split (' ');
                        string[] args = new string[6];
                        args[0] = commandSplitted[0];

                        for (int i = 1; i < commandSplitted.Length; i++) {
                            args[i] = commandSplitted[i];
                        }

                        for (int i = 1; i < commandSplitted.Length; i++) {
                            if (commandSplitted[i].IndexOf ("Arg" + i) > -1) {
                                args[i] = commandSplitted[i].Replace ("Arg" + i + "=", "");
                            }
                        }

                        var columnIndex = Array.IndexOf (columnHedaers, "Command");
                        ICell targetCell = targetRow.CreateCell (columnIndex);

                        for (int i = 0; i < args.Length; i++) {
                            if (args[i] != null) {
                                targetCell = targetRow.CreateCell (columnIndex + i);
                                targetCell.SetCellValue (args[i]);
                            }
                        }

                    }

                     
                    
                }*/
            }

            // Excelを保存
            using (fs = File.Create (excelPath)) {
                writeBook.Write (fs);
                Debug.Log ("Done:" + textPath);
            }
        }

        private void SetPageCtrlInputBr (IRow targetRow) {
            var columnIndex = Array.IndexOf (columnHedaers, "PageCtrl");
            var pageCtrlCell = targetRow.CreateCell (columnIndex);
            pageCtrlCell.SetCellValue ("InputBr");
        }

        private bool IsNewLineTag (string line) {
            return (line.IndexOf ("[p]") > -1 || line.IndexOf ("@p") > -1);
        }
        private bool IsNewLineTag (IEnumerable<string> lines) {
            return lines.Any (l => IsNewLineTag (l));
        }

        void SetColumnHeader (ISheet sheet) {
            var row = sheet.CreateRow (0);

            for (int i = 0; i < columnHedaers.Length; i++) {
                var targetCell = row.CreateCell (i);
                targetCell.SetCellValue (columnHedaers[i]);
            }
        }

        bool isOneByteChar (char ch) {
            var length = System.Text.Encoding.GetEncoding ("UTF-8").GetByteCount (new char[] { ch });
            if (length == 1)
                return true;

            return false;
        }

    }
}
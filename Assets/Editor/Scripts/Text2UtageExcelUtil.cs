using System;
using System.Collections;
using System.Collections.Generic;
using NPOI.SS.UserModel;
using UnityEngine;

namespace Text2Utage {
    public class Util {
        public enum ColumnEnum {
            Command = 0,
            Arg1,
            Arg2,
            Arg3,
            Arg4,
            Arg5,
            Arg6,
            Text,
            PageCtrl,
            Voice,
            WindowType,
            English
        }

        public static bool isOneByteChar (char ch) {
            var length = System.Text.Encoding.GetEncoding ("UTF-8").GetByteCount (new char[] { ch });
            return length == 1;
        }

        public static bool isOneByteChar (string str) {
            byte[] byteData = System.Text.Encoding.GetEncoding (932).GetBytes (str);
            return byteData.Length == str.Length;
        }
    }

    public class Sheet {
        private ISheet targetSheet;
        private bool novelMode;

        private int rowCount;

        public Sheet (ISheet seet, bool novel) {
            targetSheet = seet;
            novelMode = novel;
            SetColumnHeader ();
            rowCount = 1;
        }

        void SetColumnHeader () {
            var row = targetSheet.CreateRow (0);

            string[] array = Enum.GetNames (typeof (Util.ColumnEnum));
            for (int i = 0; i < array.Length; i++) {
                string name = array[i];
                var targetCell = row.CreateCell (i);
                targetCell.SetCellValue (name);
                Console.WriteLine (name);
            }
        }

        private bool IsArgTextInCommand (string[] commands) {
            for (int i = 1; i < commands.Length; i++) {
                if (commands[i].ToLower ().Contains ("arg")) return true;
            }
            return false;
        }

        public void SetLine (string line) {
            if (line == "") { return; }

            var firstChar = line[0];
            var targetRow = new Row (targetSheet.CreateRow (rowCount));

            //コメントの場合
            if (firstChar == ';' || firstChar == '/') {
                return;
            }

            //文章の場合
            if (!Util.isOneByteChar (line) && firstChar != '[' && firstChar != '@') {
                targetRow.SetText (line, novelMode);
                rowCount++;
                return;
            }

            //改行コマンドの場合、1つ前のPageCtrlを削除
            if (targetRow.IsNewLineTag (line)) {
                if (rowCount == 1) return;

                var beforeRow = targetSheet.GetRow (rowCount - 1);
                var pageCtrlCell = beforeRow.CreateCell ((int) Util.ColumnEnum.PageCtrl);
                pageCtrlCell.SetCellValue ("");
                return;
            }

            //ここまで来たらコマンド行とみなし、整形
            var commandText = line.Trim ('[', ']', '@').Replace (" =", "=").Replace ("= ", "=").Replace ("  ", " ");

            string[] splittedCommandText = commandText.Split (' ');
            string[] commands = new string[6];
            commands[0] = splittedCommandText[0];

            if (IsArgTextInCommand (splittedCommandText)) {
                //Arg表記ありの場合、arg指定に合わせる
                for (int i = 1; i < splittedCommandText.Length; i++) {
                    for (int argNum = 1; argNum < 6; argNum++) {
                        if (splittedCommandText[i].ToLower ().Contains ("arg" + argNum)) {
                            commands[argNum] = splittedCommandText[i].Replace ("arg" + i + "=", "");
                        }
                    }
                }
            } else {
                //Arg表記なしの場合、順に入れる
                for (int i = 1; i < splittedCommandText.Length; i++) {
                    if (splittedCommandText[i] == "") continue;
                    commands[i] = splittedCommandText[i];
                }
            }
            targetRow.SetCommand (commands);

            rowCount++;
        }
    }

    public class Row {
        private IRow targetRow;

        public Row (IRow row) {
            targetRow = row;
        }

        private void SetPageCtrlInputBr (IRow targetRow) {
            var pageCtrlCell = targetRow.CreateCell ((int) Util.ColumnEnum.PageCtrl);
            pageCtrlCell.SetCellValue ("InputBr");
        }

        public bool IsNewLineTag (string line) {
            return (line.IndexOf ("[p]") > -1 || line.IndexOf ("@p") > -1);
        }

        public void SetText (string text, bool novelMode = false) {
            var cell = targetRow.CreateCell ((int) Util.ColumnEnum.Text);

            if (novelMode && !IsNewLineTag (text)) {
                SetPageCtrlInputBr (targetRow);
            }

            cell.SetCellValue (text.Replace ("[p]", "").Replace ("@p", ""));
        }

        public void SetCommand (string[] commands) {
            var commandCell = targetRow.CreateCell ((int) Util.ColumnEnum.Command);
            commandCell.SetCellValue (commands[0]);

            for (int i = 1; i < commands.Length; i++) {
                if (!string.IsNullOrEmpty (commands[i])) {
                    var targetCell = targetRow.CreateCell (i);
                    targetCell.SetCellValue (commands[i]);
                }
            }
        }

    }

}
using System;
using System.Collections;
using System.Collections.Generic;
using System.Text.RegularExpressions;
using NPOI.SS.UserModel;
using UnityEngine;

namespace Text2Utage
{
	public class Util
	{
		public enum ColumnEnum
		{
			Command = 0,
			Arg1,
			Arg2,
			Arg3,
			Arg4,
			Arg5,
			Arg6,
			WaitType,
			Text,
			PageCtrl,
			Voice,
			WindowType,
			English
		}

		public static bool isOneByteChar(char ch)
		{
			var length = System.Text.Encoding.GetEncoding("UTF-8").GetByteCount(new char[] { ch });
			return length == 1;
		}

		public static bool isOneByteChar(string str)
		{
			byte[] byteData = System.Text.Encoding.GetEncoding(932).GetBytes(str);
			return byteData.Length == str.Length;
		}

		public static bool isCommand(string str)
		{
			return str[0] == '@' || str[0] == '[';
		}
	}

	public class Sheet
	{
		private ISheet targetSheet;
		private bool novelMode;

		private int rowCount;

		private bool isPreviousCharactorCommand;

		public Sheet(ISheet seet, bool novel)
		{
			targetSheet = seet;
			novelMode = novel;
			SetColumnHeader();
			rowCount = 1;
		}

		void SetColumnHeader()
		{
			var row = targetSheet.CreateRow(0);

			var columnArray = Enum.GetNames(typeof(Util.ColumnEnum));
			for (int i = 0; i < columnArray.Length; i++)
			{
				string name = columnArray[i];
				var targetCell = row.CreateCell(i);
				targetCell.SetCellValue(name);
				Console.WriteLine(name);
			}
		}

		private bool IsArgTextInCommand(string[] commands)
		{
			for (int i = 0; i < commands.Length; i++)
			{
				return commands[i].ToLower().Contains("arg");
			}
			return false;
		}

		public void SetLine(string line)
		{
			if (line == "") { return; }

			var firstChar = line[0];
			Row targetRow;

			if (isPreviousCharactorCommand)
			{
				targetRow = new Row(targetSheet.GetRow(targetSheet.LastRowNum));
				isPreviousCharactorCommand = false;
			}
			else
			{
				targetRow = new Row(targetSheet.CreateRow(rowCount));
			}

			//コメントの場合
			if (firstChar == ';' || firstChar == '/')
			{
				return;
			}

			//コマンドではない
			if (!Util.isCommand(line))
			{
				//キャラクターコマンドの場合
				if (firstChar == '【')
				{
					var argsSpiritedText = line.Replace("【", "").Replace("】", "").Replace("  ", " ").Split(' ');
					argsSpiritedText[0] = "arg1" + argsSpiritedText[0];

					var args = GetArgContentArray(argsSpiritedText);
					//名前
					args[0] = Regex.Match(line, "(?<=【).*?(?=】)").Value;

					targetRow.SetCommand("", args);
					isPreviousCharactorCommand = true;
					return;
				}

				targetRow.SetText(line, novelMode);
				rowCount++;
				return;
			}

			//改行コマンドの場合、1つ前のPageCtrlを削除
			if (targetRow.IsNewLineTag(line))
			{
				if (rowCount == 1) return;

				var beforeRow = targetSheet.GetRow(rowCount - 1);
				var pageCtrlCell = beforeRow.CreateCell((int)Util.ColumnEnum.PageCtrl);
				pageCtrlCell.SetCellValue("");
				return;
			}
			// 入力待ちの場合、1つ前のPageCtrlをInputにする
			if (targetRow.IsClickWaitTag(line))
			{
				if (rowCount == 1) return;

				var beforeRow = targetSheet.GetRow(rowCount - 1);
				var pageCtrlCell = beforeRow.CreateCell((int)Util.ColumnEnum.PageCtrl);
				pageCtrlCell.SetCellValue("Input");
				return;
			}

			//ここまで来たらコマンド行とみなし、整形
			var commandText = line.Trim('[', ']', '@').Replace(" =", "=").Replace("= ", "=").Replace("  ", " ");

			string[] splittedCommandText = commandText.Split(' ');
			string[] argSplittedText = new string[splittedCommandText.Length - 1];

			Array.Copy(splittedCommandText, 1, argSplittedText, 0, splittedCommandText.Length - 1);


			if (splittedCommandText[0] == "Tween")
			{

			}

			targetRow.SetCommand(splittedCommandText[0], GetArgContentArray(argSplittedText));

			rowCount++;
		}

		private string[] GetArgContentArray(string[] splittedArgText)
		{
			string[] args = new string[8];
			if (splittedArgText == null)
			{
				return args;
			}

			if (IsArgTextInCommand(splittedArgText))
			{
				//Arg表記ありの場合、arg指定に合わせる
				for (int i = 0; i < splittedArgText.Length; i++)
				{
					for (int argNum = 0; argNum < 6; argNum++)
					{
						if (splittedArgText[i].ToLower().Contains("arg" + (argNum + 1)))
						{
							args[argNum] = splittedArgText[i]
								.Replace("arg" + (argNum + 1) + "=", "")
								.Replace("Arg" + (argNum + 1) + "=", "")
								.Replace("ARG" + (argNum + 1) + "=", "");

							//Tween用
							if (args[argNum].Contains("{"))
							{
								args[argNum] = args[argNum]
									.Replace("{", "")
									.Replace("}", "")
									.Replace(",", " ");
							}
						}

						if (splittedArgText[i].ToLower().Contains(Util.ColumnEnum.WaitType.ToString().ToLower()))
						{
							args[6] = splittedArgText[i].Replace(Util.ColumnEnum.WaitType.ToString() + "=", "");
						}
						if (splittedArgText[i].ToLower().Contains(Util.ColumnEnum.Text.ToString().ToLower()))
						{
							args[7] = splittedArgText[i].Replace(Util.ColumnEnum.Text.ToString() + "=", "");
						}
					}
				}
			}
			else
			{
				//Arg表記なしの場合、順に入れる
				for (int i = 1; i < splittedArgText.Length; i++)
				{
					if (splittedArgText[i] == "") continue;
					args[i] = splittedArgText[i];
				}
			}



			return args;
		}
	}



	public class Row
	{
		private IRow targetRow;

		public Row(IRow row)
		{
			targetRow = row;
		}

		private void SetPageCtrlInputBr(IRow targetRow)
		{
			var pageCtrlCell = targetRow.CreateCell((int)Util.ColumnEnum.PageCtrl);
			pageCtrlCell.SetCellValue("InputBr");
		}

		public bool IsNewLineTag(string line)
		{
			return (line.IndexOf("[p]") > -1 || line.IndexOf("@p") > -1);
		}

		public bool IsClickWaitTag(string line)
		{
			return (line.IndexOf("[l]]") > -1 || line.IndexOf("@l") > -1);
		}

		public void SetText(string text, bool novelMode = false)
		{
			var cell = targetRow.CreateCell((int)Util.ColumnEnum.Text);

			if (novelMode && !IsNewLineTag(text))
			{
				SetPageCtrlInputBr(targetRow);
			}

			cell.SetCellValue(text.Replace("[p]", "").Replace("@p", ""));
		}

		public void SetCommand(string command, string[] args)
		{
			var commandCell = targetRow.CreateCell((int)Util.ColumnEnum.Command);

			//宴の仕様上、キャラクターだけを表示する命令はコマンドが空白
			if (command.ToLower() == "characteron")
			{
				command = "";
			}

			commandCell.SetCellValue(command);

			for (int i = 0; i < args.Length; i++)
			{
				if (!string.IsNullOrEmpty(args[i]))
				{
					var targetCell = targetRow.CreateCell(i + 1);
					targetCell.SetCellValue(args[i]);
				}
			}
		}

	}

}
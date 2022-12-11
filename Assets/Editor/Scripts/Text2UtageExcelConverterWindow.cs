using System;
using System.Collections;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Security.Cryptography;
using System.Text.RegularExpressions;
using NPOI.HSSF.UserModel;
using NPOI.SS.UserModel;
using NPOI.XSSF.UserModel;
using Text2Utage;
using UnityEditor;
using UnityEngine;

namespace Utage
{
	public class Text2UtageExcelConverterWindow : EditorWindow
	{
		public const string MenuToolRoot = "Tools/Utage/";
		static public Text2UtageExcelConverter Converter
		{
			get
			{
				return converter;
			}
			set
			{
				if (converter != value)
				{
					converter = value;
				}
			}
		}
		//プロジェクトデータ
		static Text2UtageExcelConverter converter;

		[MenuItem(MenuToolRoot + "Text2ExcelConverter")]
		static void CreateNewProject()
		{
			GetWindow(typeof(Text2UtageExcelConverterWindow), false, "Text2ExcelConverter");
		}

		void OnGUI()
		{

			Text2UtageExcelConverter project = EditorGUILayout.ObjectField("", Converter, typeof(Text2UtageExcelConverter), false) as Text2UtageExcelConverter;

			if (project != Converter)
			{
				Converter = project;
			}

			if (GUILayout.Button("Convert", GUILayout.Width(180)))
			{
				if (Converter == null)
				{
					Debug.Log("ConverterのScriptableObjectが指定されていません");
					return;
				}

				EditorUtility.ClearProgressBar();
				Convert();
			}
		}

		public static void ConvertFromFile(Text2UtageExcelConverter target)
		{
			float doneExcelFileCount = 0f;
			float convertCount = (float)target.TextCount;
			foreach (var chapter in target.ExcelChapters)
			{
				foreach (var text in chapter.TextFiles)
				{
					if (!target.IsUpdate(text))
					{
						continue;
					}
					Text2Excel(chapter.ExcelFile, text, target);
					doneExcelFileCount++;
					EditorUtility.DisplayProgressBar("Convert", $"progress {doneExcelFileCount} / {convertCount}", doneExcelFileCount / convertCount);
				}
			}

			EditorUtility.ClearProgressBar();

			if (target.CheckUpdateFile)
			{
				EditorUtility.SetDirty(target);
			}

			if (doneExcelFileCount == 0f)
			{
				Debug.Log("更新されたファイルが存在しないと判定されました");
			}
		}

		private void Convert()
		{
			ConvertFromFile(converter);
		}

		private static void Text2Excel(UnityEngine.Object excelFile, UnityEngine.Object textFile, Text2UtageExcelConverter target)
		{
			var excelPath = AssetDatabase.GetAssetPath(excelFile);
			var textPath = AssetDatabase.GetAssetPath(textFile);

			Debug.Log("Convert :" + textPath);

			IWorkbook writeBook;
			FileStream fs = null;

			// 書込用Excelファイルの読込
			using (fs = File.OpenRead(excelPath))
			{
				writeBook = WorkbookFactory.Create(fs);
			}
			var sheetName = Path.GetFileNameWithoutExtension(textPath);
			var targetSheet = writeBook.GetSheet(sheetName);

			if (targetSheet != null)
			{
				writeBook.RemoveSheetAt(writeBook.GetSheetIndex(sheetName));
			}

			targetSheet = writeBook.CreateSheet(Path.GetFileNameWithoutExtension(textPath));


			var sheet = new Text2Utage.Sheet(targetSheet, target.NovelMode);

			using (System.IO.StreamReader sr = new System.IO.StreamReader(textPath))
			{
				var textData = sr.ReadToEnd();

				string[] temp = textData.Replace("\r", "").Split('\n');
				List<string> splitted = new List<string>();
				foreach (var line in temp)
				{
					if (string.IsNullOrEmpty(line) || line[0] == ';' || line[0] == '/')
					{
						continue;
					}
					splitted.Add(line);
				}

				for (int i = 0; i < splitted.Count; i++)
				{
					if (splitted[i] == "") continue;
					//ノベルモードで[p]が存在しない場合、自動でいれる
					if (target.NovelMode && !IsNewLineTag(splitted.Skip(i)) && !Util.isCommand(splitted[i]))
					{
						int rowSum = 0;
						for (int search = i; search < splitted.Count - 1; search++)
						{
							var nowRow = Mathf.FloorToInt(splitted[search].Length / target.TextMatrixXY.x) + 1;
							if (nowRow > target.TextMatrixXY.y)
							{
								splitted[search] = splitted[search] + "[p]";
								break;
							}

							rowSum += nowRow;

							var nextRow = Mathf.FloorToInt(splitted[search + 1].Length / target.TextMatrixXY.x) + 1;

							//Debug.Log(splitted[search] + splitted[search].Length + "Row" + rowSum);
							if ((rowSum + nextRow) > target.TextMatrixXY.y)
							{
								splitted[search] = splitted[search] + "[p]";
								//                               Debug.Log ("index : "+ search + " " +splitted[search]);
								break;
							}
						}
					}

					var line = splitted[i];
					var test = line;
					sheet.SetLine(line);
				}

			}
			if (target.CheckUpdateFile)
			{
				target.AddConverted(textFile);
			}

			// Excelを保存
			using (fs = File.Create(excelPath))
			{
				writeBook.Write(fs);
				Debug.Log("Done:" + textPath);
			}
		}

		private static bool IsNewLineTag(string line)
		{
			return (line.IndexOf("[p]") > -1 || line.IndexOf("@p") > -1);
		}
		private static bool IsNewLineTag(IEnumerable<string> lines)
		{
			return lines.Any(l => IsNewLineTag(l));
		}
	}
}
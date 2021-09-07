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

		private void Convert()
		{
			float doneExcelFileCount = 0f;
			float convertCount = (float)Converter.TextCount;
			foreach (var chapter in converter.ExcelChapters)
			{
				foreach (var text in chapter.TextFiles)
				{
					if (!converter.IsUpdate(text))
					{
						continue;
					}
					Text2Excel(chapter.ExcelFile, text);
					doneExcelFileCount++;
					EditorUtility.DisplayProgressBar("Convert", $"progress {doneExcelFileCount} / {convertCount}", doneExcelFileCount / convertCount);
				}
			}

			EditorUtility.ClearProgressBar();

			if (converter.CheckUpdateFile)
			{
				EditorUtility.SetDirty(converter);
			}
			else
			{
				Debug.Log("更新されたファイルが存在しないと判定されました");
			}
		}

		private void Text2Excel(UnityEngine.Object excelFile, UnityEngine.Object textFile)
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

			var sheet = new Text2Utage.Sheet(targetSheet, Converter.NovelMode);

			using (System.IO.StreamReader sr = new System.IO.StreamReader(textPath))
			{
				var textData = sr.ReadToEnd();
				string[] splitted = textData.Replace("\r", "").Split('\n');

				for (int i = 0; i < splitted.Length; i++)
				{
					if (splitted[i] == "") continue;
					//ノベルモードで[p]が存在しない場合、自動でいれる
					if (converter.NovelMode && !IsNewLineTag(splitted.Skip(i)))
					{
						int rowSum = 0;
						for (int search = i; search < splitted.Length - 1; search++)
						{

							var nowRow = Mathf.FloorToInt(splitted[search].Length / converter.TextMatrixXY.x) + 1;
							if (nowRow > converter.TextMatrixXY.y)
							{
								splitted[search] = splitted[search] + "[p]";
								break;
							}

							rowSum += nowRow;

							var nextRow = Mathf.FloorToInt(splitted[search + 1].Length / converter.TextMatrixXY.x) + 1;

							//Debug.Log(splitted[search] + splitted[search].Length + "Row" + rowSum);
							if ((rowSum + nextRow) > converter.TextMatrixXY.y)
							{
								splitted[search] = splitted[search] + "[p]";
								//                               Debug.Log ("index : "+ search + " " +splitted[search]);
								break;
							}
						}
					}

					var line = splitted[i];
					sheet.SetLine(line);
				}

			}
			if (converter.CheckUpdateFile)
			{
				converter.AddConverted(textFile);
			}

			// Excelを保存
			using (fs = File.Create(excelPath))
			{
				writeBook.Write(fs);
				Debug.Log("Done:" + textPath);
			}
		}

		private bool IsNewLineTag(string line)
		{
			return (line.IndexOf("[p]") > -1 || line.IndexOf("@p") > -1);
		}
		private bool IsNewLineTag(IEnumerable<string> lines)
		{
			return lines.Any(l => IsNewLineTag(l));
		}
	}
}
using System;
using System.Collections;
using System.Collections.Generic;
using System.Linq;
using System.IO;
using System.Security.Cryptography;
using NPOI.SS.UserModel;
using UnityEditor;
using UnityEngine;

namespace Utage
{
	public class Text2UtageExcelConverter : ScriptableObject
	{

		[System.Serializable]
		public class ExcelChapter : System.Object
		{
			[SerializeField]
			UnityEngine.Object _excelFile;

			[SerializeField]
			List<UnityEngine.Object> _textFiles;

			public UnityEngine.Object ExcelFile
			{
				get { return _excelFile; }
#if UNITY_EDITOR
				set { _excelFile = value; }
#endif
			}

			public List<UnityEngine.Object> TextFiles
			{
				get { return _textFiles; }
#if UNITY_EDITOR
				set { _textFiles = value; }
#endif
			}
		}

		[System.Serializable]
		public class ConvertFile
		{
			[SerializeField]
			public UnityEngine.Object TextFile;

			[SerializeField]
			private string previouMd5Sum;

			public bool IsUpdate()
			{
				return previouMd5Sum != MD5Sum();
			}

			public void SetMD5()
			{
				previouMd5Sum = MD5Sum();
			}

			private string MD5Sum()
			{
				var fullName = AssetDatabase.GetAssetPath(TextFile);
				FileStream fs = new FileStream(fullName, FileMode.Open);
				string md5sum = BitConverter.ToString(MD5.Create().ComputeHash(fs)).ToLower().Replace("-", "");
				fs.Close();
				return md5sum;
			}
		}

		//		[HideInInspector]
		public List<ConvertFile> converted = new List<ConvertFile>();

		public void Convert()
		{
			Text2UtageExcelConverterWindow.ConvertFromFile(this);
		}

		public void AddConverted(UnityEngine.Object textFile)
		{
			var convertedFlie = converted.Find(c => c.TextFile == textFile);
			if (convertedFlie == null)
			{
				convertedFlie = new ConvertFile();
				convertedFlie.TextFile = textFile;
				converted.Add(convertedFlie);
			}
			convertedFlie.SetMD5();
		}

		public bool IsUpdate(UnityEngine.Object textFile)
		{
			if (!checkUpdateFile)
			{
				return true;
			}

			var convertedFlie = converted.Find(c => c.TextFile == textFile);
			if (convertedFlie != null)
			{
				return convertedFlie.IsUpdate();
			}

			return true;
		}


		[SerializeField]
		List<ExcelChapter> _excelChapters;

		public List<ExcelChapter> ExcelChapters
		{
			get
			{

				return _excelChapters;
			}
#if UNITY_EDITOR
			set { _excelChapters = value; }
#endif
		}

		public bool NovelMode
		{
			get
			{
				return novelMode;
			}

			set
			{
				novelMode = value;
			}
		}

		public int TextCount
		{
			get
			{
				return _excelChapters.SelectMany(c => c.TextFiles).Count();
			}
		}

		public Vector2Int TextMatrixXY
		{
			get
			{
				return textMatrixXY;
			}
		}

		[SerializeField, HeaderAttribute("改行クリック待ちで前の文章を消さない")]
		private bool novelMode = false;

		[SerializeField, HeaderAttribute("ノベル時の一画面の列文字数と行数")]
		private Vector2Int textMatrixXY;


		[SerializeField, HeaderAttribute("テキストに更新がない場合は変換しない")]
		private bool checkUpdateFile;

		public bool CheckUpdateFile { get => checkUpdateFile; }

	}

}
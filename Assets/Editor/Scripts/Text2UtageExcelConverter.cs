using System;
using System.Collections;
using System.Collections.Generic;
using NPOI.SS.UserModel;
using UnityEditor;
using UnityEngine;

namespace Utage {
    public class Text2UtageExcelConverter : ScriptableObject {

        [System.Serializable]
        public class ExcelChapter : System.Object {
            [SerializeField]
            UnityEngine.Object _excelFile;

            [SerializeField]
            List<UnityEngine.Object> _textFiles;

            public UnityEngine.Object ExcelFile {
                get { return _excelFile; }
#if UNITY_EDITOR
                set { _excelFile = value; }
#endif
            }

            public List<UnityEngine.Object> TextFiles {
                get { return _textFiles; }
#if UNITY_EDITOR
                set { _textFiles = value; }
#endif
            }

        }

        [SerializeField]
        List<ExcelChapter> _excelChapters;

        public List<ExcelChapter> ExcelChapters {
            get {

                return _excelChapters;
            }
#if UNITY_EDITOR
            set { _excelChapters = value; }
#endif
        }

        public bool NovelMode {
            get {
                return novelMode;
            }

            set {
                novelMode = value;
            }
        }

        public Vector2Int TextMatrixXY {
            get {
                return textMatrixXY;
            }
        }

        [SerializeField, HeaderAttribute ("改行クリック待ちで前の文章を消さない")]
        private bool novelMode = false;

        [SerializeField, HeaderAttribute ("ノベル時の一画面の列文字数と行数")]
        private Vector2Int textMatrixXY;

    }

}